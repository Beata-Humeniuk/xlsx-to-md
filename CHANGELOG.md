# Changelog

This file lists user-visible changes to Excel & CSV to Markdown. The project follows
[Semantic Versioning](https://semver.org/).

## [0.2.1] - 2026-10-05

### Fixed

- Polish letters shown as “�” in workbooks whose parts are written in
  Windows-1250 (`encoding="windows-1250"`, used by some exporters): the
  declared encoding is now honoured, as Excel does. CSV files that mix UTF-8
  and Windows-1250 lines are read line by line. Characters already lost in
  the file itself are pointed out above the grid.

- Excel's temporary lock files (`~$name.xlsx`, `~$name.xlsm`, present while
  a workbook is open in Excel) are no longer offered in the Explorer menu.
  Opened anyway, they explain what they are and offer to open the workbook
  itself instead of reporting "not an Excel workbook".

## [0.2.0] - 2026-10-05

### Added

- CSV and TSV files, with the same editor, conditions and saved filters as
  workbooks. The separator and encoding (UTF-8, UTF-16, Windows-1250) are
  recognized automatically; Excel's `sep=` line is honoured.

### Changed

- The extension is called **Excel & CSV to Markdown** (the identifier stays
  `xlsx-to-md`).

## [0.1.0] - 2026-10-05

### Added

- **Export to Markdown…** for `.xlsx`, `.xlsm`, `.xltx` and `.xltm` files: an
  editor with the sheet grid, data choices and a live Markdown preview.
- Excel-like selection: drag, Ctrl/⌘ + drag for more ranges, Shift + click,
  whole columns and rows, double-click for the block of data, the name box,
  Excel tables and "all data on the sheet" in one click.
- Columns chosen and ordered by header name, row conditions (=, ≠, contains,
  starts/ends with, >, ≥, <, ≤, empty) with value lists taken from the data,
  sorting, hidden and empty rows and columns, merged cells.
- Saved filters, available in every project: values can be fixed or given each
  time the filter is used; the filter finds its data again by sheet, Excel
  table and header names when the workbook changes.
- Output as a GFM table or a section per row, with an optional heading.
- English and Polish interface.
