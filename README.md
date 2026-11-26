# 🎞️ Movie_Database

A companion to [Movie_ID](https://github.com/spamfam/Movie_ID). Movie_ID puts IMDb IDs into your movie folder names; **Movie_Database copies the release year and ID from those folder names into your Excel spreadsheet**, filling in only what's missing.

No spreadsheet yet? It can build a starter one from your folders.

---

## Features

- 📊 **Fills blanks only** — rows that already have a year and an ID are skipped
- 🔎 **Matches on words, not just letters** — `Blade II` = `Blade 2`, `Shrek the Third` = `Shrek 3`, and `Spectre` is found inside `James Bond - Spectre`
- ✅ **You confirm uncertain matches** — a review window shows candidates side by side; back up, pick a folder by hand, or skip
- 🧾 **Summary before saving** — nothing is written until you press **Confirm all changes**; **Cancel without saving** discards everything
- ⚠️ **Never overwrites your data** — if the sheet disagrees with a folder, your value is kept and the cell is highlighted yellow with a comment
- 🗂️ **Collections and shorts** — searches inside `(No IMDB)` collection folders and untagged folders like `The Hobbit Collection`, at any depth
- 💾 **Automatic backup** — the original is copied to `yourfile.backup-<timestamp>.xlsx` before saving
- 🧪 **Dry run by default** — preview everything first
- 🆕 **Starter spreadsheet** — a small link creates a Movie Title / Year / IMDb ID sheet pre-filled from your folders

---

## Folder naming format

Movie_Database reads folder names in the format Movie_ID produces:

```
Movie Title (Year) {imdb-tt0068646}
Movie Title (Year) {tmdb-3137}
```

| Tag | Meaning |
|---|---|
| `(Year)` | Release year, taken from the last `(YYYY)` in the name |
| `{imdb-tt…}` | IMDb ID |
| `{tmdb-…}` | TMDB ID, for titles that are not on IMDb |
| `(No IMDB)` | Collection folder (home movies, shorts, boxed sets) — searched, but never matched itself |
| `{edition-…}` and other `{…}` tags | Ignored |

---

## How a row is handled

| Row in the spreadsheet | What happens |
|---|---|
| Year **and** an ID (`imdb-tt…` or `tmdb-…`) | Skipped |
| Year and `NA` in the ID column | Skipped (the title has no database ID) |
| Year or ID missing | The matching folder is found and the blanks are filled from its name |
| Value disagrees with the folder | Your value is kept, cell highlighted yellow, comment added |
| Titles in both the movie and short-title columns | The short title is used |
| No similar folder | Listed at the end of the log |

New IDs are written in the style already used in your ID column (`imdb-tt0068646`, `tt0068646` or digits only); `tmdb-…` IDs are always written as `tmdb-…`.

---

## Installation

```bash
git clone https://github.com/spamfam/Movie_Database.git
cd Movie_Database
pip install openpyxl
```

`tkinter` is also required (included in the standard Python installer from [python.org](https://www.python.org/downloads/) — choose **Customize installation** and keep **tcl/tk and IDLE** ticked).

---

## Usage

```bash
python movie_database.py
```

1. **Movies folder** — browse to your movies directory
2. **Excel file** — browse to your spreadsheet (`.xlsx` / `.xlsm`); the sheet, header row and columns are detected and can be changed in the dropdowns
3. Leave **Dry run** ticked for the first pass, then press **Start**
4. Review uncertain matches, check the summary, and press **Confirm all changes**
5. Untick **Dry run** and repeat to actually write the changes

Close the workbook in Excel before saving.

### The review window

- Each uncertain row shows your spreadsheet entry beside the best matching folders with a similarity score
- **📁 Pick a folder myself…** — choose any folder; its name is read immediately
- **← Back** — return to the previous row; your earlier choice is remembered
- **Cancel without saving** — available on every page; nothing is written
- The final page lists everything that will change; **✔ Confirm all changes** saves

### Create a starter spreadsheet

After choosing a movies folder, use the small *"No spreadsheet yet?"* link under the Excel file box. It saves a sheet with **Movie Title**, **Year** and **IMDb ID** columns, one row per folder, and loads it straight into the app.

### Options

- **Search inside every folder** — also descend into movie folders (by default only top-level and collection folders are searched)
- **Also show rows with no match in the review window** — lets you pick a folder by hand for rows nothing matched

---

## Limitations

- Excel files are edited with [openpyxl](https://openpyxl.readthedocs.io/): images, charts and text boxes in the workbook are **not preserved** on save (formulas, array formulas, formatting and comments are). The automatic backup is your safety net.
- `.xls` (old Excel) files are not supported.
- Only top-level and collection folders are searched unless you tick *Search inside every folder*.

---

## Roadmap

- [ ] 
- [ ] 

---

## Author

[spamfam](https://github.com/spamfam)
