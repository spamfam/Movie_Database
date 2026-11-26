# Changelog

All notable changes to Movie_Database are documented here.

---

## [Unreleased]

---

## [0.1.0] — Initial release

### Added
- **GUI** with folder and Excel pickers, sheet / header-row / column dropdowns (auto-detected from the header row), progress bar and live log
- **Fill-the-blanks logic** — reads year and ID from `Title (Year) {imdb-tt…}` / `{tmdb-…}` folder names and fills only missing cells
- **Skip rules** — rows with a year plus an IMDb ID, TMDB ID or `NA` are skipped
- **Short-title column** — optional second title column; a short title takes priority over the movie title
- **Word-based title matching** — roman numerals, spelled-out and ordinal numbers, `Part` / `Chapter` / `Vol` / `Episode` labels, dropped articles, titles inside longer titles (`Spectre` in `James Bond - Spectre`), typo tolerance, and a trailing `1` ignored
- **Review window** — candidate folders with similarity scores, **Pick a folder myself**, **Back**, and a final summary page with **Confirm all changes** / **Cancel without saving**
- **Disagreement flags** — conflicting values are kept, highlighted yellow and given a cell comment
- **Collection search** — descends into `(No IMDB)` and untagged collection folders at any depth; optional *Search inside every folder*
- **ID style detection** — new IDs match the style already in the sheet; `tmdb-…` IDs supported
- **Safety** — dry run on by default, timestamped backup before every save, clear message if the workbook is open in Excel
- **Starter spreadsheet** — creates a pre-filled Movie Title / Year / IMDb ID workbook from a movies folder (bold header, frozen top row, sized columns)

---
