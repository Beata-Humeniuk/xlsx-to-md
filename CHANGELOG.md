# Changelog

This file lists user-visible changes to Excel to Markdown. The project follows
[Semantic Versioning](https://semver.org/).

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
