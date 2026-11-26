#!/usr/bin/env python3
"""
Movie Database  -  companion to Movie ID
========================================
Fills in missing release years and IMDb IDs in an Excel spreadsheet by reading
them from your movie folder names, e.g.

    Through the Mirror (1936) {imdb-tt0028378}
    ^ title              ^ year   ^ IMDb ID

For every row with a movie (or short film) title (if a row has both, the
short title in its own column wins):
  * year AND ID already filled   -> skipped. "ID" can be an IMDb ID
                                    (imdb-tt0337563), a TMDB ID (tmdb-3137),
                                    or NA for titles that are in no database
  * year filled but ID is NA     -> skipped (so is any row with year + any ID)
  * either one missing           -> the matching folder is found and the blanks
                                    are filled in from its name
  * a value that disagrees with the folder is never overwritten; the cell is
    highlighted yellow and given a comment instead

Folder search:
  * every top-level folder in the movies directory
  * plus everything inside collection folders - those tagged "(No IMDB)" and
    any folder with no (year) or {imdb-...} tag - at any depth
  * optionally every folder at every depth

Titles are matched on words, not just letters: "Blade II" = "Blade 2",
"Shrek the Third" = "Shrek 3", and "Spectre" is found inside
"James Bond - Spectre".

No spreadsheet yet? Once a movies folder is chosen, a small link in the Files
box can create a starter spreadsheet (Movie Title / Year / IMDb ID) pre-filled
from the folder names.

Before running:
    pip install openpyxl

Usage:
    python movie_database.py
"""

import os
import re
import sys
import shutil
import difflib
import threading
import traceback
import unicodedata
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox
from datetime import datetime

try:
    import openpyxl
    from openpyxl.comments import Comment
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl is not installed.")
    print("Fix it by running:  py -m pip install openpyxl")
    sys.exit(1)


# ─── Settings ─────────────────────────────────────────────────────────────────

AUTO_ACCEPT    = 0.97   # title similarity needed to accept a match without asking
REVIEW_MIN     = 0.60   # below this a folder is not even offered as a candidate
MAX_CANDIDATES = 5
POOL_LIMIT     = 150    # most folders compared in detail per row (speed guard)

FLAG_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

YEAR_RE    = re.compile(r"\((\d{4})\)")
ID_RE      = re.compile(r"\{imdb-(tt\d+)\}", re.IGNORECASE)
NO_IMDB_RE = re.compile(r"\(No IMDB\)", re.IGNORECASE)
TMDB_RE    = re.compile(r"\{tmdb-(\d+)\}", re.IGNORECASE)
TAG_RE     = re.compile(r"\{[^}]*\}")        # any {...} tag, incl. {edition-...}
NA_RE      = re.compile(r"^\s*n\s*/?\s*a\s*$", re.IGNORECASE)   # NA, N/A, n/a

# Sequel numbers are compared as digits, so "Blade II", "Blade 2" and
# "Blade Two" all become "blade 2". ("i" and "x" are left alone on purpose.)
ROMAN_NUMERALS = {"ii": "2", "iii": "3", "iv": "4", "v": "5", "vi": "6", "vii": "7",
                  "viii": "8", "ix": "9", "xi": "11", "xii": "12", "xiii": "13",
                  "xiv": "14", "xv": "15", "xvi": "16", "xvii": "17", "xviii": "18",
                  "xix": "19", "xx": "20"}
NUMBER_WORDS = {"two": "2", "second": "2", "three": "3", "third": "3", "four": "4",
                "fourth": "4", "five": "5", "fifth": "5", "six": "6", "sixth": "6",
                "seven": "7", "seventh": "7", "eight": "8", "eighth": "8",
                "nine": "9", "ninth": "9", "ten": "10", "tenth": "10",
                "eleven": "11", "twelve": "12"}
SECTION_WORDS = {"part", "pt", "vol", "volume", "chapter", "episode"}   # "Part 3" == "3"
ARTICLES      = {"the", "a", "an"}


# ─── Text / value helpers ─────────────────────────────────────────────────────

def strip_accents(text):
    """Amelie == Amélie, but leave non-Latin scripts (e.g. Japanese) alone."""
    out = []
    for ch in unicodedata.normalize("NFC", text):
        decomposed = unicodedata.normalize("NFKD", ch)
        out.append(decomposed[0] if decomposed and ord(decomposed[0]) < 128 else ch)
    return "".join(out)


def normalize_title(text):
    """
    Reduce a title to a comparable form. All of these become 'karate kid 3':
        'The Karate Kid Part III' / 'Karate Kid, Part 3' / 'Karate Kid Third'
    Articles are dropped, punctuation removed, and sequel numbers become digits.
    """
    text = strip_accents(str(text)).lower().replace("&", " and ")
    text = text.replace("'", "").replace("\u2019", "").strip()
    m = re.match(r"^(.*?),\s*(the|a|an)$", text)          # 'Godfather, The'
    if m:
        text = m.group(2) + " " + m.group(1)
    text = re.sub(r"[^\w\s]", " ", text)
    tokens = text.split()

    converted = []
    for i, tok in enumerate(tokens):
        prev = tokens[i - 1] if i else ""
        if tok in ROMAN_NUMERALS:
            tok = ROMAN_NUMERALS[tok]
        elif tok in NUMBER_WORDS:
            tok = NUMBER_WORDS[tok]
        elif tok in ("one", "i") and prev in SECTION_WORDS:   # "Part One", "Episode I"
            tok = "1"
        converted.append(tok)

    converted = [t for t in converted if t not in ARTICLES]
    result = []
    for i, tok in enumerate(converted):
        if tok in SECTION_WORDS and i + 1 < len(converted) and converted[i + 1].isdigit():
            continue                                       # "part 3" -> "3"
        result.append(tok)
    return " ".join(result)


def strip_trailing_one(norm):
    """'young and dangerous 1' -> 'young and dangerous' (film 1 often has no number)."""
    tokens = norm.split()
    if len(tokens) > 1 and tokens[-1] == "1":
        return " ".join(tokens[:-1])
    return None


def sequel_numbers(tokens):
    """The 1-2 digit numbers in a title (sequel markers, not years like 2049)."""
    return {t for t in tokens if t.isdigit() and len(t) <= 2}


def normalize_id(digits):
    """'68646' / '0068646' -> 'tt0068646'"""
    return "tt" + str(int(digits)).zfill(7)


def parse_year_value(value):
    """Pull a 4-digit year out of a cell value (int, float, or text)."""
    if value is None:
        return None
    m = re.search(r"(?<!\d)(1[89]\d{2}|20\d{2})(?!\d)", str(value))
    return m.group(1) if m else None


def parse_id_value(value):
    """
    Pull a database ID out of a cell value:
      imdb-tt0337563 / tt0337563 / 337563  -> 'tt0337563'
      tmdb-3137                            -> 'tmdb-3137'
    """
    if value is None:
        return None
    text = str(value)
    m = re.search(r"tmdb[-_\s]?(\d+)", text, re.IGNORECASE)   # check TMDB first
    if m:
        return "tmdb-" + str(int(m.group(1)))
    m = re.search(r"(?:tt)?(\d{5,9})", text, re.IGNORECASE)
    return normalize_id(m.group(1)) if m else None


def is_na_value(value):
    """True for cells deliberately marked NA / N/A (title has no database ID)."""
    return value is not None and bool(NA_RE.match(str(value)))


def same_system(a, b):
    """IMDb IDs only compare with IMDb IDs, TMDB with TMDB."""
    return a.startswith("tmdb-") == b.startswith("tmdb-")


def detect_id_style(ws, cols, header_row):
    """
    Look at the IDs already in the sheet so new ones match.
      'imdb-tt' -> imdb-tt0337563   (default)
      'tt'      -> tt0337563
      'digits'  -> 0337563
    """
    checked = 0
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=cols["imdb_id"]).value
        if v is None or not str(v).strip():
            continue
        text = str(v).strip()
        if re.match(r"imdb-tt\d+", text, re.IGNORECASE):
            return "imdb-tt"
        if re.match(r"tt\d+", text, re.IGNORECASE):
            return "tt"
        if re.fullmatch(r"\d{5,9}", text):
            return "digits"
        checked += 1
        if checked >= 50:
            break
    return "imdb-tt"


def format_id(db_id, style):
    if db_id.startswith("tmdb-"):
        return db_id
    if style == "tt":
        return db_id
    if style == "digits":
        return db_id[2:]
    return "imdb-" + db_id


def parse_folder(name):
    """
    'Through the Mirror (1936) {imdb-tt0028378}'
        -> title 'Through the Mirror', year '1936', db_id 'tt0028378'
    (a {tmdb-3137} tag is understood too, giving db_id 'tmdb-3137')
    """
    imdb_match = ID_RE.search(name)
    tmdb_match = TMDB_RE.search(name)
    years = list(YEAR_RE.finditer(name))
    title = name[:years[-1].start()] if years else TAG_RE.sub("", name)
    if imdb_match:
        db_id = normalize_id(imdb_match.group(1)[2:])
    elif tmdb_match:
        db_id = "tmdb-" + str(int(tmdb_match.group(1)))
    else:
        db_id = None
    return {
        "title":   title.strip(),
        "year":    years[-1].group(1) if years else None,
        "db_id":   db_id,
    }


# ─── Folder index ─────────────────────────────────────────────────────────────

def collect_candidates(root, search_all=False):
    """
    Walk the movies directory and return every folder that carries a (year) or
    {imdb-tt...} / {tmdb-...} tag. We descend into collection folders - ones
    tagged "(No IMDB)" or with no year/ID tag at all, e.g. "The Hobbit
    Collection" - but not into movie folders themselves, unless search_all is on.
    """
    cands = []

    def walk(path):
        try:
            names = sorted(os.listdir(path))
        except OSError:
            return
        for name in names:
            full = os.path.join(path, name)
            if not os.path.isdir(full) or os.path.islink(full):
                continue
            info = parse_folder(name)
            is_movie = bool(info["year"] or info["db_id"])
            if is_movie:
                cands.append({
                    "name":    name,
                    "path":    full,
                    "parent":  "" if path == root else os.path.relpath(path, root),
                    "title":   info["title"],
                    "norm":    normalize_title(info["title"]),
                    "year":    info["year"],
                    "db_id":   info["db_id"],
                })
            if search_all or NO_IMDB_RE.search(name) or not is_movie:
                walk(full)

    walk(root)
    return [c for c in cands if c["norm"]]


def candidate_from_path(path, root=None):
    """
    Build a candidate from a folder the user picked by hand.
    Returns None if its name has no (year) or ID tag to copy.
    """
    path = os.path.normpath(path)
    name = os.path.basename(path)
    info = parse_folder(name)
    if not (info["year"] or info["db_id"]):
        return None
    parent = os.path.dirname(path)
    try:
        rel = os.path.relpath(parent, root) if root else parent
    except ValueError:                      # different drive on Windows
        rel = parent
    if rel == ".":
        rel = ""
    norm = normalize_title(info["title"])
    return {"name": name, "path": path, "parent": rel, "title": info["title"],
            "norm": norm, "tokens": norm.split(),
            "year": info["year"], "db_id": info["db_id"]}


# ─── Matching ─────────────────────────────────────────────────────────────────

def _find_inside(short, long_):
    """
    Look for the words of `short` inside `long_` as one consecutive run, ignoring
    any sequel numbers in `long_` that `short` does not have - so
    "Star Wars A New Hope" is found inside "Star Wars Episode 4 A New Hope".

    Returns None if the run is not there. Otherwise returns (followed, whole):
      followed - the run is immediately followed by such a number, i.e. a
                 different installment ("Ghostbusters" inside "Ghostbusters 2")
      whole    - nothing but number labels was left over
    """
    own = sequel_numbers(short)

    def is_label(tok):
        return tok.isdigit() and len(tok) <= 2 and tok not in own

    kept = [(i, t) for i, t in enumerate(long_) if not is_label(t)]
    words = [t for _, t in kept]
    n = len(short)
    for start in range(len(words) - n + 1):
        if words[start:start + n] == short:
            last = kept[start + n - 1][0]            # where the run ends in long_
            followed = last + 1 < len(long_) and is_label(long_[last + 1])
            return followed, len(words) == n
    return None


def title_score(e_tokens, c_tokens):
    """
    Similarity (0..1) between a spreadsheet title and a folder title, both
    already normalized and split into words.
      * identical                       -> 1.0
      * identical apart from a sequel-number label in the middle or front
        ("Star Wars New Hope" / "Star Wars Episode 4 New Hope")  -> 0.98
      * letter similarity (difflib)     -> typos and spelling variants
      * one title inside the other      -> 0.70-0.95  ("Spectre" inside
                                           "James Bond Spectre", "Empire Strikes
                                           Back" inside "Star Wars Episode 5 ...")
      * the run is followed by a sequel number, or the two titles carry
        different sequel numbers        -> minus 0.30 ("Ghostbusters" is a poor
                                           match for "Ghostbusters 2")
    """
    e, c = " ".join(e_tokens), " ".join(c_tokens)
    if e == c:
        return 1.0

    matcher = difflib.SequenceMatcher(None, e, c)
    score = 0.0
    if matcher.real_quick_ratio() >= REVIEW_MIN and matcher.quick_ratio() >= REVIEW_MIN:
        score = matcher.ratio()

    short, long_ = (e_tokens, c_tokens) if len(e_tokens) <= len(c_tokens) else (c_tokens, e_tokens)
    substantial = len(short) >= 2 or len(short[0]) >= 5
    found = _find_inside(short, long_) if substantial else None

    if found is not None:
        followed, whole = found
        coverage = len(" ".join(short)) / len(" ".join(long_))
        if whole and not followed:
            score = max(score, 0.98)
        else:
            score = max(score, 0.70 + 0.25 * coverage)
        if followed:
            score -= 0.30
    elif sequel_numbers(e_tokens) != sequel_numbers(c_tokens):
        score -= 0.30
    return score


class FolderIndex:
    """Lookup tables over the folder candidates so every row is matched quickly."""

    def __init__(self, cands):
        self.by_norm  = {}   # normalized title            -> [candidates]
        self.by_alias = {}   # title without a trailing 1  -> [candidates]
        self.by_token = {}   # significant word            -> [candidates]
        self.vocab_by_first = {}   # first letter -> words (for typo lookups)
        for c in cands:
            c["tokens"] = c["norm"].split()
            self.by_norm.setdefault(c["norm"], []).append(c)
            alias = strip_trailing_one(c["norm"])
            if alias:
                self.by_alias.setdefault(alias, []).append(c)
            for tok in set(c["tokens"]):
                if len(tok) >= 3 and not tok.isdigit():
                    self.by_token.setdefault(tok, []).append(c)
        for tok in self.by_token:
            self.vocab_by_first.setdefault(tok[0], []).append(tok)


def find_matches(row, index):
    """
    Return up to MAX_CANDIDATES (score, candidate) pairs for a spreadsheet row,
    best first. Exact (normalized) title matches are tried first; only when
    there is none do we look for partial and similar titles.
    """
    clean = YEAR_RE.sub("", TAG_RE.sub("", row["title"]))
    norm = normalize_title(clean)
    if not norm:
        return []
    tokens = norm.split()

    best = {}   # id(candidate) -> (score, candidate)

    def offer(score, cand):
        if id(cand) not in best or score > best[id(cand)][0]:
            best[id(cand)] = (score, cand)

    for cand in index.by_norm.get(norm, []):
        offer(1.0, cand)
    alias = strip_trailing_one(norm)                    # sheet 'x 1' ~ folder 'x'
    if alias:
        for cand in index.by_norm.get(alias, []):
            offer(0.99, cand)
    for cand in index.by_alias.get(norm, []):           # sheet 'x' ~ folder 'x 1'
        offer(0.99, cand)

    if not best:
        # Shortlist folders that share a word with the row (folders sharing the
        # most words first). A word found in no folder may be a typo, so also
        # try the closest known words that start with the same letter.
        pool, shared = {}, {}

        def add_word(word):
            for cand in index.by_token.get(word, []):
                shared[id(cand)] = shared.get(id(cand), 0) + 1
                pool[id(cand)] = cand

        for tok in set(tokens):
            if tok in index.by_token:
                add_word(tok)
            elif len(tok) >= 4 and not tok.isdigit():
                bucket = index.vocab_by_first.get(tok[0], [])
                for near in difflib.get_close_matches(tok, bucket, n=3, cutoff=0.75):
                    add_word(near)

        for key in sorted(pool, key=lambda k: shared[k], reverse=True)[:POOL_LIMIT]:
            cand = pool[key]
            score = title_score(tokens, cand["tokens"])
            if score >= REVIEW_MIN:
                offer(score, cand)

    scored = list(best.values())
    want_short = row["kind"] == "short"

    def sort_key(item):
        score, cand = item
        year_hit = bool(row["year"]) and cand["year"] == row["year"]
        # shorts prefer folders found inside collections, movies prefer top-level
        place_hit = bool(cand["parent"]) == want_short
        return (score, year_hit, place_hit)

    scored.sort(key=sort_key, reverse=True)
    return scored[:MAX_CANDIDATES]


def conflicts(row, cand):
    """Messages describing where the spreadsheet disagrees with the folder."""
    msgs = []
    if row["year"] and cand["year"] and row["year"] != cand["year"]:
        msgs.append(f"Year in spreadsheet ({row['year']}) differs from folder ({cand['year']})")
    if (row["db_id"] and cand["db_id"] and same_system(row["db_id"], cand["db_id"])
            and row["db_id"] != cand["db_id"]):
        msgs.append(f"ID in spreadsheet ({row['db_id']}) differs from folder ({cand['db_id']})")
    return msgs


def classify(row, scored):
    """
    Decide what to do with a row:
      'notfound' - no folder is similar enough
      'auto'     - one unambiguous exact title match with no disagreements
      'review'   - anything else; the user picks
    """
    if not scored:
        return "notfound", None
    top_score = scored[0][0]
    top = [s for s in scored if s[0] == top_score]
    if row["year"]:
        same_year = [s for s in top if s[1]["year"] == row["year"]]
        if len(same_year) == 1:
            top = same_year
    if top_score >= AUTO_ACCEPT and len(top) == 1 and not conflicts(row, top[0][1]):
        return "auto", top[0][1]
    return "review", scored


# ─── Spreadsheet reading / writing ────────────────────────────────────────────

def read_rows(ws, cols, header_row):
    """
    Read every row that has a title. If a row has both a short title and a movie
    title, the short title wins. Titles are converted to text because some are
    numbers (e.g. the film '300').
    """
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        title, kind = None, None
        for key in ("short", "movie"):
            col = cols.get(key)
            if not col:
                continue
            v = ws.cell(row=r, column=col).value
            if v is not None and str(v).strip():
                title, kind = str(v).strip(), key
                break
        if title is None:
            continue
        raw_year = ws.cell(row=r, column=cols["year"]).value
        raw_id   = ws.cell(row=r, column=cols["imdb_id"]).value
        rows.append({
            "row": r, "title": title, "kind": kind,
            "year": parse_year_value(raw_year), "db_id": parse_id_value(raw_id),
            "id_na": is_na_value(raw_id),
            "raw_year": raw_year, "raw_id": raw_id,
        })
    return rows


def scan_rows(rows, cands, log, progress_fn=None, stop_flag=None, review_unmatched=False):
    """
    Phase 1 (read-only): decide what to do with each row.
    Returns a plan dict, or None if the user pressed Stop.
    """
    index = FolderIndex(cands)

    plan  = {"complete": 0, "auto": [], "review": [], "notfound": []}
    total = len(rows)

    for i, row in enumerate(rows, 1):
        if stop_flag and stop_flag.is_set():
            return None
        if progress_fn:
            progress_fn(i, total)

        tag = "short" if row["kind"] == "short" else "movie"
        if row["year"] and (row["db_id"] or row["id_na"]):
            plan["complete"] += 1
            why = "year filled, ID marked NA" if row["id_na"] else "year + ID already filled"
            log(f"[SKIP]    Row {row['row']}: {row['title']}  ({why})")
            continue

        kind, payload = classify(row, find_matches(row, index))
        if kind == "auto":
            plan["auto"].append((row, payload))
            log(f"[MATCH]   Row {row['row']} ({tag}): {row['title']}  ->  {payload['name']}")
        elif kind == "review":
            plan["review"].append({"row": row, "scored": payload})
            log(f"[REVIEW]  Row {row['row']} ({tag}): {row['title']}  ->  needs your confirmation")
        elif review_unmatched:
            plan["review"].append({"row": row, "scored": []})
            log(f"[REVIEW]  Row {row['row']} ({tag}): {row['title']}  ->  no automatic match - pick a folder by hand")
        else:
            plan["notfound"].append(row)
            log(f"[NONE]    Row {row['row']} ({tag}): {row['title']}  ->  no matching folder")

    return plan


def _was(raw):
    return f"   (replacing '{raw}')" if raw is not None and str(raw).strip() else ""


def set_cell(ws, row, col, value, log):
    """Write a value; merged cells can't be written, so report instead of crashing."""
    try:
        ws.cell(row=row, column=col).value = value
        return True
    except AttributeError:
        log("          ✗ cell is part of a merged range - could not write")
        return False


def flag_cell(ws, row, col, message):
    """Highlight a cell yellow and attach a comment explaining the disagreement."""
    try:
        cell = ws.cell(row=row, column=col)
        cell.fill = FLAG_FILL
        cell.comment = Comment(message, "Movie Database")
    except AttributeError:
        pass


def apply_match(ws, cols, row, cand, id_style, do_write, log):
    """
    Fill blank cells from the folder; never overwrite a real value.
    Returns (years_filled, ids_filled, flags_raised).
    """
    r = row["row"]
    years = ids = flags = 0

    # ── Year ──
    if row["year"] is None:
        if cand["year"]:
            ok = set_cell(ws, r, cols["year"], int(cand["year"]), log) if do_write else True
            if ok:
                log(f"          + year  {cand['year']}{_was(row['raw_year'])}")
                years += 1
        else:
            log("          ! folder name has no (year) to copy")
    elif cand["year"] and cand["year"] != row["year"]:
        if do_write:
            flag_cell(ws, r, cols["year"], f"Folder '{cand['name']}' says {cand['year']}")
        log(f"          ⚠ year in sheet ({row['year']}) differs from folder ({cand['year']}) - kept, highlighted")
        flags += 1

    # ── ID (an NA cell is left alone; IMDb and TMDB IDs never overwrite each other) ──
    if row["id_na"]:
        pass
    elif row["db_id"] is None:
        if cand["db_id"]:
            new_value = format_id(cand["db_id"], id_style)
            ok = set_cell(ws, r, cols["imdb_id"], new_value, log) if do_write else True
            if ok:
                log(f"          + ID    {new_value}{_was(row['raw_id'])}")
                ids += 1
        else:
            log("          ! folder has no {imdb-tt...} or {tmdb-...} tag yet (run Movie ID on it first)")
    elif (cand["db_id"] and same_system(row["db_id"], cand["db_id"])
          and cand["db_id"] != row["db_id"]):
        if do_write:
            flag_cell(ws, r, cols["imdb_id"], f"Folder '{cand['name']}' says {cand['db_id']}")
        log(f"          ⚠ ID in sheet ({row['db_id']}) differs from folder ({cand['db_id']}) - kept, highlighted")
        flags += 1

    return years, ids, flags


def apply_plan(ws, cols, plan, decisions, id_style, do_write, log):
    """Phase 2: write the auto matches and the user's confirmed choices."""
    t = {"auto": 0, "confirmed": 0, "review_skipped": 0,
         "unreviewed": max(0, len(plan["review"]) - len(decisions)),
         "years": 0, "ids": 0, "flags": 0}

    def run(label, row, cand):
        log(f"[{label}] Row {row['row']}: {row['title']}")
        log(f"          folder: {cand['name']}" +
            (f"   (in {cand['parent']})" if cand["parent"] else ""))
        y, i, f = apply_match(ws, cols, row, cand, id_style, do_write, log)
        t["years"] += y
        t["ids"]   += i
        t["flags"] += f
        log("")

    for row, cand in plan["auto"]:
        t["auto"] += 1
        run("FILL     ", row, cand)

    for row, cand in decisions:
        if cand is None:
            t["review_skipped"] += 1
            log(f"[SKIP]    Row {row['row']}: {row['title']}  (skipped in review)")
        else:
            t["confirmed"] += 1
            run("CONFIRMED", row, cand)

    return t


def make_backup(path):
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    base, ext = os.path.splitext(path)
    backup = f"{base}.backup-{stamp}{ext}"
    shutil.copy2(path, backup)
    return backup


def write_starter_workbook(cands, path):
    """
    Create a basic spreadsheet - Movie Title / Year / IMDb ID - pre-filled from
    folder candidates (see collect_candidates). IMDb and TMDB IDs both work.
    The headers use wording the column guesser recognizes, so the file can be
    used straight away. Returns the number of movie rows written.
    """
    rows = sorted(cands, key=lambda c: (c["norm"], c["year"] or ""))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Movies"
    ws.append(["Movie Title", "Year", "IMDb ID"])
    for cand in rows:
        ws.append([
            cand["title"],
            int(cand["year"]) if cand["year"] else None,
            format_id(cand["db_id"], "imdb-tt") if cand["db_id"] else None,
        ])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    longest = max((len(c["title"]) for c in rows), default=0)
    ws.column_dimensions["A"].width = min(max(longest + 2, 20), 60)
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 18

    wb.save(path)
    return len(rows)


# ─── Review window ────────────────────────────────────────────────────────────

MANUAL_CHOICE = 99   # radio value meaning "the folder I picked by hand"


class ReviewQueue(tk.Toplevel):
    """
    Walks through every uncertain row, then shows a final summary of ALL pending
    changes. Nothing is written until you press "Confirm all changes" on that
    summary page; "Cancel without saving" (or closing the window) discards
    everything.

      on_confirm(decisions) - decisions: list of (row, chosen_candidate_or_None)
      on_cancel()           - nothing is to be saved

    Not modal, so the main window's log stays scrollable while you review.
    """

    def __init__(self, parent, items, on_confirm, on_cancel, start_dir=None,
                 auto_matches=None, do_write=True):
        super().__init__(parent)
        self.title("Confirm Matches")
        self.transient(parent)          # stays on top of the main window
        self.minsize(780, 680)

        self.items        = items
        self.on_confirm   = on_confirm
        self.on_cancel    = on_cancel
        self.start_dir    = start_dir
        self.auto_matches = auto_matches or []
        self.do_write     = do_write
        self.index        = 0
        self.choices      = [None] * len(items)   # radio choice remembered per row
        self.manual       = {}                    # row index -> folder picked by hand
        self._closed      = False
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self._build_review_ui()
        self._build_summary_ui()
        if self.items:
            self.review_frame.pack(fill="both", expand=True)
            self._load_item()
        else:
            self._show_summary()

    # ── Per-row review page ──

    def _build_review_ui(self):
        rf = self.review_frame = tk.Frame(self)

        self.header = tk.Label(rf, text="", font=("Helvetica", 12, "bold"))
        self.header.pack(fill="x", padx=12, pady=6)

        compare = tk.Frame(rf, relief="groove", bd=1)
        compare.pack(fill="x", padx=12, pady=4)
        compare.columnconfigure(0, weight=1)
        compare.columnconfigure(1, weight=1)

        tk.Label(compare, text="Spreadsheet row", font=("Helvetica", 9, "bold"),
                 fg="#555").grid(row=0, column=0, sticky="w", padx=8, pady=(6, 0))
        tk.Label(compare, text="Selected folder", font=("Helvetica", 9, "bold"),
                 fg="#555").grid(row=0, column=1, sticky="w", padx=8, pady=(6, 0))

        self.row_label = tk.Label(compare, anchor="w", justify="left", wraplength=340)
        self.row_label.grid(row=1, column=0, sticky="w", padx=8, pady=(0, 8))
        self.cand_label = tk.Label(compare, anchor="w", justify="left",
                                   wraplength=340, fg="#1a6eb5")
        self.cand_label.grid(row=1, column=1, sticky="w", padx=8, pady=(0, 8))

        self.warn_label = tk.Label(rf, text="", fg="#b00020", anchor="w",
                                   justify="left", wraplength=720)
        self.warn_label.pack(fill="x", padx=12)

        tk.Label(rf, text="Choose the matching folder:", anchor="w").pack(
            fill="x", padx=12, pady=(8, 0))

        radio_frame = tk.Frame(rf)
        radio_frame.pack(fill="x", padx=12)
        self.choice_var = tk.IntVar(value=0)
        self.radios = []
        for i in range(MAX_CANDIDATES):
            rb = tk.Radiobutton(radio_frame, variable=self.choice_var, value=i,
                                anchor="w", justify="left", wraplength=720,
                                command=self._on_radio)
            rb.pack(fill="x", pady=1)
            self.radios.append(rb)

        self.manual_radio = tk.Radiobutton(
            radio_frame, variable=self.choice_var, value=MANUAL_CHOICE,
            anchor="w", justify="left", wraplength=720, command=self._on_radio)
        self.manual_radio.pack(fill="x", pady=(6, 0))

        tk.Radiobutton(radio_frame, text="⏭  None of these - skip this row",
                       variable=self.choice_var, value=-1, anchor="w",
                       command=self._on_radio).pack(fill="x", pady=(6, 0))

        tk.Button(rf, text="📁  Pick a folder myself…",
                  command=self._browse_manual).pack(anchor="w", padx=12, pady=(10, 0))

        nav = tk.Frame(rf)
        nav.pack(fill="x", padx=12, pady=12)
        tk.Button(nav, text="Cancel without saving", command=self._cancel,
                  fg="#b00020").pack(side="left")
        self.back_btn = tk.Button(nav, text="← Back", command=self._back, width=10)
        self.back_btn.pack(side="left", padx=(8, 0))
        self.next_btn = tk.Button(nav, text="Confirm & Next →", command=self._confirm,
                                  bg="#4CAF50", fg="white", width=26)
        self.next_btn.pack(side="right")
        self.progress_lbl = tk.Label(nav, text="")
        self.progress_lbl.pack(side="right", padx=12)

    def _load_item(self):
        item   = self.items[self.index]
        row    = item["row"]
        scored = item["scored"]
        kind   = "short film" if row["kind"] == "short" else "movie"
        last   = self.index == len(self.items) - 1

        self.header.config(text=f"Review {self.index + 1} of {len(self.items)}")
        self.progress_lbl.config(text=f"{self.index + 1} / {len(self.items)}")
        self.back_btn.config(state="normal" if self.index > 0 else "disabled")
        self.next_btn.config(text="Confirm & See Summary →" if last else "Confirm & Next →")
        self.row_label.config(
            text=(f"Row {row['row']}  ({kind})\n{row['title']}\n"
                  f"Year: {row['year'] or '—'}    ID: {row['db_id'] or ('NA' if row['id_na'] else '—')}"))

        for i, rb in enumerate(self.radios):
            if i < len(scored):
                score, cand = scored[i]
                where = f"     [in: {cand['parent']}]" if cand["parent"] else ""
                rb.config(text=f"{score:.0%}   {cand['name']}{where}", state="normal")
            else:
                rb.config(text="—", state="disabled")

        self._refresh_manual_radio()
        saved = self.choices[self.index]            # restore if we came back to this row
        self.choice_var.set((0 if scored else -1) if saved is None else saved)
        self._on_radio()

    def _refresh_manual_radio(self):
        cand = self.manual.get(self.index)
        if cand:
            where = f"     [in: {cand['parent']}]" if cand["parent"] else ""
            self.manual_radio.config(text=f"📁   Picked by you:   {cand['name']}{where}",
                                     state="normal")
        else:
            self.manual_radio.config(
                text="📁   (nothing picked yet - use the button below)", state="disabled")

    def _candidate_for(self, index, choice):
        """The candidate a given radio choice means for a given row, or None (skip)."""
        if choice == MANUAL_CHOICE:
            return self.manual.get(index)
        scored = self.items[index]["scored"]
        if choice is not None and 0 <= choice < len(scored):
            return scored[choice][1]
        return None

    def _chosen(self):
        return self._candidate_for(self.index, self.choice_var.get())

    def _browse_manual(self):
        """Let the user pick the right folder; its name is parsed immediately."""
        options = {"parent": self, "title": "Select the matching movie folder"}
        if self.start_dir and os.path.isdir(self.start_dir):
            options["initialdir"] = self.start_dir
        path = filedialog.askdirectory(**options)
        if not path:
            return
        cand = candidate_from_path(path, self.start_dir)
        if cand is None:
            messagebox.showwarning(
                "Nothing to copy",
                f"'{os.path.basename(os.path.normpath(path))}' has no (year) or "
                "{imdb-tt...} tag in its name, so there is nothing to copy from it.",
                parent=self)
            return
        self.manual[self.index] = cand
        self._refresh_manual_radio()
        self.choice_var.set(MANUAL_CHOICE)
        self._on_radio()

    def _on_radio(self):
        cand = self._chosen()
        if cand is None:
            self.cand_label.config(text="(this row will be skipped)")
            self.warn_label.config(text="")
            return
        text = f"{cand['name']}\nYear: {cand['year'] or '—'}    ID: {cand['db_id'] or '—'}"
        if cand["parent"]:
            text += f"\nin: {cand['parent']}"
        self.cand_label.config(text=text)

        msgs = conflicts(self.items[self.index]["row"], cand)
        if msgs:
            self.warn_label.config(
                text="⚠ " + "; ".join(msgs) +
                     ". Your spreadsheet value will be kept and highlighted.")
        else:
            self.warn_label.config(text="")

    def _confirm(self):
        """Remember this row's choice and move on (to the summary after the last row)."""
        self.choices[self.index] = self.choice_var.get()
        self.index += 1
        if self.index < len(self.items):
            self._load_item()
        else:
            self._show_summary()

    def _back(self):
        """Return to the previous row; its earlier choice is restored."""
        if self.index == 0:
            return
        self.choices[self.index] = self.choice_var.get()
        self.index -= 1
        self._load_item()

    # ── Final summary page ──

    def _build_summary_ui(self):
        sf = self.summary_frame = tk.Frame(self)

        tk.Label(sf, text="Review complete - check everything before saving",
                 font=("Helvetica", 12, "bold"), anchor="w").pack(fill="x", padx=12, pady=(10, 4))
        self.summary_counts = tk.Label(sf, text="", anchor="w", justify="left")
        self.summary_counts.pack(fill="x", padx=12)

        self.summary_box = scrolledtext.ScrolledText(sf, height=22, font=("Courier", 9),
                                                     wrap="word", state="disabled")
        self.summary_box.pack(fill="both", expand=True, padx=12, pady=8)

        nav = tk.Frame(sf)
        nav.pack(fill="x", padx=12, pady=12)
        tk.Button(nav, text="Cancel without saving", command=self._cancel,
                  fg="#b00020").pack(side="left")
        self.summary_back_btn = tk.Button(nav, text="← Back", command=self._back_from_summary,
                                          width=10)
        self.summary_back_btn.pack(side="left", padx=(8, 0))
        label = "✔  Confirm all changes" if self.do_write else "Finish dry run"
        tk.Button(nav, text=label, command=self._confirm_all,
                  bg="#4CAF50", fg="white", width=26).pack(side="right")

    def _decisions(self):
        return [(item["row"], self._candidate_for(i, self.choices[i]))
                for i, item in enumerate(self.items)]

    @staticmethod
    def _summary_line(row, cand):
        text = f"Row {row['row']}: {row['title']}  →  "
        text += cand["name"] if cand else "(skipped)"
        if cand and conflicts(row, cand):
            text += "   ⚠ disagrees with your sheet - sheet value kept, highlighted"
        return text

    def _show_summary(self):
        self.review_frame.pack_forget()
        decisions = self._decisions()
        confirmed = [(r, c) for r, c in decisions if c is not None]
        skipped   = [(r, c) for r, c in decisions if c is None]

        lines = []
        if self.auto_matches:
            lines.append(f"MATCHED AUTOMATICALLY ({len(self.auto_matches)})")
            lines += [self._summary_line(r, c) for r, c in self.auto_matches]
            lines.append("")
        if confirmed:
            lines.append(f"CONFIRMED BY YOU ({len(confirmed)})")
            lines += [self._summary_line(r, c) for r, c in confirmed]
            lines.append("")
        if skipped:
            lines.append(f"SKIPPED ({len(skipped)})")
            lines += [self._summary_line(r, c) for r, c in skipped]

        mode = "" if self.do_write else "DRY RUN - nothing will be written.   "
        self.summary_counts.config(
            text=(f"{mode}{len(self.auto_matches)} matched automatically,  "
                  f"{len(confirmed)} confirmed by you,  {len(skipped)} skipped."))
        self.summary_back_btn.config(state="normal" if self.items else "disabled")

        self.summary_box.config(state="normal")
        self.summary_box.delete("1.0", "end")
        self.summary_box.insert("end", "\n".join(lines))
        self.summary_box.config(state="disabled")
        self.summary_frame.pack(fill="both", expand=True)

    def _back_from_summary(self):
        if not self.items:
            return
        self.summary_frame.pack_forget()
        self.review_frame.pack(fill="both", expand=True)
        self.index = len(self.items) - 1
        self._load_item()

    # ── Finishing ──

    def _confirm_all(self):
        if self._closed:
            return
        self._closed = True
        decisions = self._decisions()
        self.destroy()
        self.on_confirm(decisions)

    def _cancel(self):
        """Discard everything (also what closing the window does)."""
        if self._closed:
            return
        something_to_lose = self.do_write and (self.auto_matches or
                                               any(ch is not None for ch in self.choices))
        if something_to_lose and not messagebox.askyesno(
                "Cancel without saving",
                "Discard everything and close?\nNothing will be written to the spreadsheet.",
                parent=self):
            return
        self._closed = True
        self.destroy()
        self.on_cancel()


# ─── Main window ──────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Movie Database")
        self.minsize(760, 700)
        self._stop_flag = threading.Event()
        self.headers = []
        self._build_ui()

    # ── UI ──

    def _build_ui(self):
        pad = dict(padx=10, pady=4)

        files = tk.LabelFrame(self, text="Files")
        files.pack(fill="x", **pad)
        files.columnconfigure(1, weight=1)

        tk.Label(files, text="Movies folder:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.folder_var = tk.StringVar()
        tk.Entry(files, textvariable=self.folder_var).grid(row=0, column=1, sticky="ew", padx=4)
        tk.Button(files, text="Browse…", command=self._browse_folder).grid(row=0, column=2, padx=6)

        tk.Label(files, text="Excel file:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.excel_var = tk.StringVar()
        tk.Entry(files, textvariable=self.excel_var).grid(row=1, column=1, sticky="ew", padx=4)
        tk.Button(files, text="Browse…", command=self._browse_excel).grid(row=1, column=2, padx=6)

        # Discreet: only needed for first-time users or as a backup, so it is a small link
        self.create_btn = tk.Button(
            files, text="No spreadsheet yet? Create one from this folder…",
            command=self._create_spreadsheet, relief="flat", bd=0, cursor="hand2",
            fg="#666666", font=("Helvetica", 8, "underline"), state="disabled")
        self.create_btn.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 4))
        self.folder_var.trace_add("write", self._update_create_btn)

        layout = tk.LabelFrame(self, text="Spreadsheet layout")
        layout.pack(fill="x", **pad)
        layout.columnconfigure(1, weight=1)
        layout.columnconfigure(3, weight=1)

        tk.Label(layout, text="Sheet:").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        self.sheet_combo = ttk.Combobox(layout, state="readonly", width=24)
        self.sheet_combo.grid(row=0, column=1, sticky="w", padx=4)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self._load_headers(guess_row=True))

        tk.Label(layout, text="Header row:").grid(row=0, column=2, sticky="e", padx=6)
        self.header_row_var = tk.StringVar(value="1")
        spin = tk.Spinbox(layout, from_=1, to=500, width=5, textvariable=self.header_row_var,
                          command=self._load_headers)
        spin.grid(row=0, column=3, sticky="w")
        spin.bind("<Return>", lambda e: self._load_headers())
        spin.bind("<FocusOut>", lambda e: self._load_headers())

        self.combos = {}
        specs = [("movie", "Movie titles column:", 1, 0),
                 ("short", "Short titles column:", 1, 2),
                 ("year",  "Release year column:", 2, 0),
                 ("imdb_id", "IMDb ID column:",     2, 2)]
        for key, label, r, c in specs:
            tk.Label(layout, text=label).grid(row=r, column=c, sticky="e" if c else "w",
                                              padx=6, pady=3)
            combo = ttk.Combobox(layout, state="readonly", width=26)
            combo.grid(row=r, column=c + 1, sticky="w", padx=4)
            self.combos[key] = combo

        opts = tk.Frame(self)
        opts.pack(fill="x", **pad)
        self.dry_run_var = tk.BooleanVar(value=True)
        tk.Checkbutton(opts, text="Dry run  (preview only - the spreadsheet is not changed)",
                       variable=self.dry_run_var).pack(anchor="w")
        self.search_all_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            opts,
            text="Search inside every folder  (default: top-level folders plus collection "
                 "folders, i.e. \"(No IMDB)\" or untagged)",
            variable=self.search_all_var).pack(anchor="w")
        self.review_unmatched_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            opts,
            text="Also show rows with no match in the review window, so I can pick a folder by hand",
            variable=self.review_unmatched_var).pack(anchor="w")

        btns = tk.Frame(self)
        btns.pack(fill="x", **pad)
        self.start_btn = tk.Button(btns, text="▶  Start", command=self._start,
                                   width=12, bg="#4CAF50", fg="white")
        self.start_btn.pack(side="left", padx=(0, 5))
        self.stop_btn = tk.Button(btns, text="⛔  Stop", command=self._stop, width=12,
                                  bg="#f44336", fg="white", state="disabled")
        self.stop_btn.pack(side="left")

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(4, 0))
        self.progress_label = tk.Label(self, text="", anchor="w")
        self.progress_label.pack(fill="x", padx=10)

        tk.Label(self, text="Log:", anchor="w").pack(fill="x", padx=10)
        self.log_box = scrolledtext.ScrolledText(self, state="disabled", height=18,
                                                 font=("Courier", 10))
        self.log_box.pack(fill="both", expand=True, padx=10, pady=5)

    # ── Thread-safe helpers ──

    def _log(self, message):
        def _write():
            self.log_box.config(state="normal")
            self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(0, _write)

    def _update_progress(self, current, total):
        def _update():
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = current
            self.progress_label.config(text=f"{current} / {total} rows checked")
        self.after(0, _update)

    def _reset_buttons(self):
        self.after(0, lambda: self.start_btn.config(state="normal"))
        self.after(0, lambda: self.stop_btn.config(state="disabled"))
        self.after(0, self._update_create_btn)

    # ── File / column pickers ──

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Select your movies folder")
        if folder:
            self.folder_var.set(folder)

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select your Excel file",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            self.excel_var.set(path)
            self._load_sheets()

    def _load_sheets(self):
        path = self.excel_var.get().strip()
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            names = wb.sheetnames
            wb.close()
        except Exception as e:
            self._log(f"⚠️  Could not open workbook: {e}")
            return
        self.sheet_combo["values"] = names
        self.sheet_combo.current(0)
        self._load_headers(guess_row=True)

    def _load_headers(self, guess_row=False):
        """Read the header row so the column dropdowns show real column names."""
        path  = self.excel_var.get().strip()
        sheet = self.sheet_combo.get()
        if not path or not sheet:
            return
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            top = list(wb[sheet].iter_rows(min_row=1, max_row=30, values_only=True))
            wb.close()
        except Exception as e:
            self._log(f"⚠️  Could not read headers: {e}")
            return

        if guess_row:
            guess = 1
            for i, vals in enumerate(top, 1):
                text = [str(v).lower() for v in vals if v is not None]
                if any("year" in t for t in text) and any("title" in t for t in text):
                    guess = i
                    break
            self.header_row_var.set(str(guess))

        try:
            header_row = max(1, int(self.header_row_var.get()))
        except ValueError:
            return
        first = top[header_row - 1] if header_row <= len(top) else ()

        self.headers, labels = [], []
        for i, v in enumerate(first, 1):
            letter = get_column_letter(i)
            text = "" if v is None else str(v).strip()
            self.headers.append(text.lower())
            labels.append(f"{letter}  -  {text}" if text else letter)

        if not labels:
            return

        def guess_col(tiers, exclude=()):
            for kw in tiers:
                for i, h in enumerate(self.headers):
                    if kw in h and not any(x in h for x in exclude):
                        return i
            return None

        picks = {
            "movie":   guess_col(["movie"]) if guess_col(["movie"]) is not None
                       else guess_col(["title"], exclude=("short", "box", "tv")),
            "short":   guess_col(["short"]),
            "year":    guess_col(["year", "release"]),
            "imdb_id": guess_col(["imdb", "id#", "id"]),
        }
        for key, combo in self.combos.items():
            values = (["(none)"] + labels) if key == "short" else labels
            combo["values"] = values
            idx = picks[key]
            if key == "short":
                combo.current(0 if idx is None else idx + 1)
            else:
                combo.current(0 if idx is None else idx)

    # ── Starter spreadsheet ──

    def _update_create_btn(self, *args):
        """The link is only usable once a real movies folder has been chosen."""
        ok = os.path.isdir(self.folder_var.get().strip())
        self.create_btn.config(state="normal" if ok else "disabled")

    def _create_spreadsheet(self):
        movies_dir = self.folder_var.get().strip()
        if not os.path.isdir(movies_dir):
            messagebox.showwarning("Movies folder", "Please choose a valid movies folder first.")
            return
        start_dir = os.path.dirname(os.path.normpath(movies_dir)) or movies_dir
        path = filedialog.asksaveasfilename(
            parent=self, title="Save the new spreadsheet as",
            defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")],
            initialdir=start_dir, initialfile="Movie Database.xlsx")
        if not path:
            return
        self.create_btn.config(state="disabled")
        self.start_btn.config(state="disabled")
        self._log(f"\nCreating a spreadsheet from {movies_dir} ...")
        threading.Thread(target=self._create_phase,
                         args=(movies_dir, path, self.search_all_var.get()),
                         daemon=True).start()

    def _create_phase(self, movies_dir, path, search_all):
        """Background thread: scan the folders and write the starter workbook."""
        log = self._log
        try:
            cands = collect_candidates(movies_dir, search_all)
            if not cands:
                log("Nothing to put in a spreadsheet - no folder names contain a (year) "
                    "or an {imdb-tt...} / {tmdb-...} tag.")
                return
            count = write_starter_workbook(cands, path)
            with_id = sum(1 for c in cands if c["db_id"])
            log(f"✓ Created {path}")
            log(f"  {count} movies  ({with_id} with an ID, {count - with_id} without)\n")
            self.after(0, lambda: self._use_new_spreadsheet(path))
        except PermissionError:
            log("✗ Could not save - is a file with that name open in Excel? "
                "Close it or choose another name.")
        except Exception as e:
            log(f"✗ Could not create the spreadsheet: {e}\n{traceback.format_exc()}")
        finally:
            self._reset_buttons()

    def _use_new_spreadsheet(self, path):
        """Load the new file into the Excel box so Start works straight away."""
        self.excel_var.set(path)
        self._load_sheets()
        self._log("The new spreadsheet is loaded in the Excel box above.\n")

    # ── Run ──

    def _start(self):
        movies_dir = self.folder_var.get().strip()
        excel      = self.excel_var.get().strip()
        if not os.path.isdir(movies_dir):
            messagebox.showwarning("Movies folder", "Please choose a valid movies folder.")
            return
        if not os.path.isfile(excel) or not excel.lower().endswith((".xlsx", ".xlsm")):
            messagebox.showwarning("Excel file", "Please choose an .xlsx or .xlsm file.")
            return
        if not self.sheet_combo.get():
            messagebox.showwarning("Sheet", "Please choose a sheet.")
            return
        try:
            header_row = max(1, int(self.header_row_var.get()))
        except ValueError:
            messagebox.showwarning("Header row", "Header row must be a number.")
            return

        picks = {k: c.current() for k, c in self.combos.items()}
        if picks["movie"] < 0 or picks["year"] < 0 or picks["imdb_id"] < 0:
            messagebox.showwarning("Columns", "Please choose the movie title, year and ID columns.")
            return

        cols = {"movie": picks["movie"] + 1,
                "short": picks["short"] if picks["short"] > 0 else None,  # 0 = (none)
                "year": picks["year"] + 1,
                "imdb_id": picks["imdb_id"] + 1}
        used = [v for v in cols.values() if v]
        if len(set(used)) != len(used):
            messagebox.showwarning("Columns", "Each field needs its own column.")
            return

        settings = {
            "movies_dir": movies_dir, "excel": excel, "sheet": self.sheet_combo.get(),
            "header_row": header_row, "cols": cols,
            "do_write": not self.dry_run_var.get(),
            "search_all": self.search_all_var.get(),
            "review_unmatched": self.review_unmatched_var.get(),
        }

        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")
        self.progress["value"] = 0
        self.progress_label.config(text="")
        self._stop_flag.clear()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.create_btn.config(state="disabled")

        threading.Thread(target=self._scan_phase, args=(settings,), daemon=True).start()

    def _stop(self):
        self._stop_flag.set()
        self.stop_btn.config(state="disabled")

    def _scan_phase(self, s):
        """Background thread: index folders, read the sheet, decide on every row."""
        log = self._log
        try:
            mode = "LIVE RUN" if s["do_write"] else "DRY RUN (nothing will be written)"
            log(f"{'=' * 60}\n  {mode}\n  Folders : {s['movies_dir']}\n  Excel   : {s['excel']}\n{'=' * 60}\n")

            log("Indexing movie folders...")
            cands = collect_candidates(s["movies_dir"], s["search_all"])
            in_collections = sum(1 for c in cands if c["parent"])
            log(f"  {len(cands)} folders with a (year) or ID tag "
                f"({in_collections} inside collection folders)\n")
            if not cands:
                log("Nothing to match against - no folder names contain a (year) or {imdb-tt...} tag.")
                self._reset_buttons()
                return

            log("Reading spreadsheet...")
            wb = openpyxl.load_workbook(s["excel"], keep_vba=s["excel"].lower().endswith(".xlsm"))
            ws = wb[s["sheet"]]
            id_style = detect_id_style(ws, s["cols"], s["header_row"])
            rows = read_rows(ws, s["cols"], s["header_row"])
            example = format_id("tt0000001", id_style)
            log(f"  {len(rows)} rows with a title; new IDs will look like: {example}\n")

            plan = scan_rows(rows, cands, log, self._update_progress, self._stop_flag,
                             s["review_unmatched"])
            if plan is None:
                log("\n⛔ Stopped by user - nothing was changed.")
                wb.close()
                self._reset_buttons()
                return

            self.after(0, lambda: self._after_scan(plan, wb, ws, id_style, s))
        except Exception as e:
            log(f"\n✗ Unexpected error: {e}\n{traceback.format_exc()}")
            self._reset_buttons()

    def _after_scan(self, plan, wb, ws, id_style, s):
        """Main thread: open the review window if needed, then apply."""
        log = self._log
        log(f"\n✅ Scan complete: {plan['complete']} already complete, {len(plan['auto'])} matched, "
            f"{len(plan['review'])} need review, {len(plan['notfound'])} not found.\n")
        # Before anything is written, always give a chance to review / confirm / cancel.
        if plan["review"] or (s["do_write"] and plan["auto"]):
            log("Opening review window...\n" if plan["review"]
                else "Opening confirmation window...\n")
            self.stop_btn.config(state="disabled")
            ReviewQueue(self, plan["review"],
                        on_confirm=lambda decisions: self._finish(plan, decisions, wb, ws, id_style, s),
                        on_cancel=lambda: self._cancelled(wb),
                        start_dir=s["movies_dir"],
                        auto_matches=plan["auto"], do_write=s["do_write"])
        else:
            self._finish(plan, [], wb, ws, id_style, s)

    def _cancelled(self, wb):
        """The user chose 'Cancel without saving': leave everything untouched."""
        self._log("\n⛔ Cancelled - nothing was saved. The spreadsheet was not changed.\n")
        try:
            wb.close()
        except Exception:
            pass
        self._reset_buttons()

    def _finish(self, plan, decisions, wb, ws, id_style, s):
        """Main thread: apply changes, save with a backup, print the summary."""
        log = self._log
        try:
            t = apply_plan(ws, s["cols"], plan, decisions, id_style, s["do_write"], log)

            if s["do_write"] and (t["years"] or t["ids"] or t["flags"]):
                backup = None
                try:
                    backup = make_backup(s["excel"])
                    wb.save(s["excel"])
                    log(f"💾 Saved {s['excel']}")
                    log(f"   Backup of the original: {backup}\n")
                except PermissionError:
                    if backup and os.path.exists(backup):
                        os.remove(backup)
                    log("✗ Could not save - is the workbook open in Excel? "
                        "Close it and run again. Nothing was changed.\n")
                except Exception as e:
                    log(f"✗ Save failed: {e}\n")

            w = "Filled" if s["do_write"] else "Would fill"
            log(f"{'=' * 60}\n  SUMMARY\n{'=' * 60}")
            log(f"  ⏭  Already complete    : {plan['complete']}")
            log(f"  ✓ Auto-matched         : {t['auto']}")
            log(f"  ✓ Confirmed by you     : {t['confirmed']}")
            log(f"  ⏭  Skipped in review    : {t['review_skipped']}")
            if t["unreviewed"]:
                log(f"  ⏭  Not reviewed (window closed early): {t['unreviewed']}")
            log(f"  + {w} years       : {t['years']}")
            log(f"  + {w} IDs         : {t['ids']}")
            log(f"  ⚠ Disagreements        : {t['flags']}"
                f"{'  (highlighted yellow in the sheet)' if s['do_write'] else ''}")
            log(f"  ✗ No matching folder   : {len(plan['notfound'])}")
            if plan["notfound"]:
                log("\n  Rows with no matching folder:")
                for row in plan["notfound"]:
                    log(f"    Row {row['row']}: {row['title']}")
            if not s["do_write"]:
                log("\n  Dry run - untick 'Dry run' and run again to write these changes.")
            log("")
        except Exception as e:
            log(f"\n✗ Unexpected error: {e}\n{traceback.format_exc()}")
        finally:
            try:
                wb.close()
            except Exception:
                pass
            self._reset_buttons()


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()
