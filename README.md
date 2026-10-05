# Excel & CSV to Markdown

Take exactly the part of an Excel workbook or CSV file you need and turn it
into Markdown — not the whole file. Pick the sheet and the cells in a live grid,
choose the columns, filter the rows, see the Markdown, export. Save the setup
as a filter and next time it is one click (plus the values you chose to give
each time).

**Works offline.** The extension makes no network requests, sends no
telemetry and never runs formulas or macros: it reads the values Excel saved.

## Open a workbook

- In the Explorer, right-click an `.xlsx` (or `.xlsm`, `.xltx`, `.xltm`),
  `.csv` or `.tsv` file and choose **Export to Markdown…**.
- Or run **Excel & CSV to Markdown: Export to Markdown…** from the Command Palette.
- Or **Open With… → Excel & CSV to Markdown** on the file.

Old `.xls` files need to be saved as `.xlsx` in Excel first.

### CSV files

A CSV file opens as a workbook with one sheet, named after the file, and
everything below works the same — selection, columns, conditions, saved
filters. The separator (`,` `;` tab `|`) and the encoding (UTF-8, UTF-16, or
Windows-1250 as Excel saves CSV on Polish Windows) are recognized from the
content, and Excel's `sep=;` first line is honoured. Numbers, dates and
`TRUE`/`FALSE` are recognized so conditions compare them as such; the text
goes into the Markdown exactly as written. A filter saved on one CSV file
works on the next one, even when columns were added or moved.

## Choose the data

The workbook opens in a grid that works like Excel:

| Do this | To |
|---|---|
| Drag over cells | select a range |
| Ctrl/⌘ + drag | add another, independent range |
| Shift + click | extend the range |
| Click a column letter or row number (and drag) | select whole columns or rows |
| Double-click a cell | select the whole block of data around it |
| Ctrl/⌘ + A | select all data on the sheet |
| Type into the name box, e.g. `A4:G120, J4:K20` | select by address |

The side panel also offers the sheet's **Excel tables** and **All data on the
sheet** in one click. Nothing is selected for you: the extension does not
assume you want every sheet or the whole used range.

When the first row of the selection holds names, they become the column
names. This is detected automatically and can be switched off.

### Columns

Tick the columns you want and drag them into the order you want. Columns are
known by their **header names**, not their letters, so the choice still works
when columns move around in the workbook.

### Row conditions

Add as many conditions as you need and keep the rows that meet all of them,
or any of them:

`=` · `≠` · contains · does not contain · starts with · ends with ·
`>` · `≥` · `<` · `≤` · is empty · is not empty

Comparisons are offered for numbers and dates. The values to choose from are
the ones in the column (with how often each occurs); you can tick several, or
type one that is not there. Matching ignores case and extra spaces, compares
numbers and dates as numbers and dates, and understands `TRUE`/`FALSE`,
`yes`/`no`, `tak`/`nie`.

### Options

- skip hidden rows, hidden columns and empty rows; leave out empty columns;
  repeat the value of merged cells;
- values as shown in Excel (number and date formats) or without formats;
- a GFM **table** or a **section per row** (the first column is the heading,
  the others a list);
- a heading: the sheet name, the filter name or your own text;
- line breaks in cells as `<br>` or spaces; numbers aligned right; the table
  lined up in the source.

Several ranges become one table when their columns match, or a table each —
or as you choose.

## Preview and export

The Markdown preview under the grid updates as you click, as Markdown source
or rendered. Rows that are left out are struck through in the grid, columns
that are left out are dimmed. Then:

- **Export .md** — save to a file (next to the workbook by default);
- **Open in editor** — a new unsaved Markdown document;
- **Copy** — to the clipboard.

## Saved filters

**Save as filter…** stores the whole setup under a name: how the data is
found, the columns, the conditions and the options. Filters are **global** —
the same filter works in any project, on the same workbook again or on
another one.

### Values fixed or given when used

Each condition's value is either:

- **Fixed** — saved with the filter, e.g. `Status ≠ DELETED`;
- **Ask when used** — chosen each time, e.g. `System = ?`.

Using a filter with no asked values applies it at once. Otherwise a small
window shows only the asked fields, each with the values that exist in this
workbook (narrowed by the fixed conditions and by what you already chose).
Leave a field empty to keep every value. The values given last time are
suggested (setting `xlsxToMd.rememberValues`).

### Finding the data again

A filter remembers the sheet, the Excel table and the header names of each
range. On another workbook — or the same one after changes — it looks for the
sheet by name, then for the sheet that has those headers, and for the range by
table name, then by header row. A range that ran to the end of the data keeps
doing so (**to the end of the data**), so rows added later are included.
If something cannot be found, a note under the toolbar says what was used
instead. Columns that are new in the workbook are listed but not added.

### Changing a filter

After applying a filter you can still change anything. The filter is marked
**modified**, and your changes apply to this export only — until you choose
**Save changes** (or **Save as new…**, or **Revert**).

### Managing filters

**Excel & CSV to Markdown: Manage Saved Filters** lists the filters with what they
do, and lets you rename or delete them. The same list is linked from the
extension's **Saved filters** settings section and from the filter list in
the editor.

## Settings

| Setting | Default | Purpose |
|---|---|---|
| `xlsxToMd.openAfterExport` | `editor` | After exporting to a file: open it in the `editor`, the `preview`, or `none`. |
| `xlsxToMd.rememberValues` | `true` | Suggest the values given last time for conditions asked when a filter is used. |

The interface follows the editor's language (English and Polish).
