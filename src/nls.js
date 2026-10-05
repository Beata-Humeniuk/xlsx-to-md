'use strict';

// UI message catalog, shared by the extension host and the webview. English
// is the base every unset key falls back to; Polish is a full translation.
// The language follows the editor's display language (vscode.env.language).
// The Markdown keeps the workbook's own words and is not translated.

// Plural keys hold arrays: English [one, many], Polish [one, few, many].
const EN = {
  'action.overwrite': 'Overwrite',
  'action.delete': 'Delete',
  'action.rename': 'Rename',
  'action.openSettings': 'Open settings',

  'dialog.openLabel': 'Open',
  'dialog.filter': 'Excel workbooks',
  'dialog.saveLabel': 'Export',
  'dialog.markdown': 'Markdown',

  'error.notXlsx': '{file} is not an Excel workbook (.xlsx).',
  'error.oldXls': '{file} is in the old .xls format. Open it in Excel and save it as .xlsx (Excel Workbook), then try again.',
  'error.failed': 'Could not read {file}: {reason}',
  'error.noSheets': '{file} has no worksheets with cells.',

  'info.exported': 'Saved {file}.',
  'info.copied': 'Markdown copied to the clipboard.',
  'info.filterSaved': 'Saved filter “{name}”.',
  'info.filterUpdated': 'Filter “{name}” updated.',
  'info.filterDeleted': 'Deleted filter “{name}”.',
  'info.noFilters': 'There are no saved filters yet. Open an Excel file, choose the data you want and save the setup as a filter.',

  'prompt.filterName': 'Name of the filter',
  'prompt.filterNamePlaceholder': 'e.g. Message specification',
  'prompt.filterNameEmpty': 'Give the filter a name.',
  'prompt.filterExists': 'A filter named “{name}” already exists. Replace it?',
  'prompt.replace': 'Replace',
  'prompt.deleteFilter': 'Delete the filter “{name}”? This cannot be undone.',
  'prompt.renameFilter': 'New name of the filter',
  'prompt.nameTaken': 'A filter named “{name}” already exists.',

  'manage.title': 'Saved filters',
  'manage.placeholder': 'Saved filters are available in every project. Use the buttons to rename or delete.',

  'summary.sheet': 'sheet {sheet}',
  'summary.columns': '{count} {columns}',
  'summary.conditions': '{count} {conditions}',
  'summary.ask': '{count} given when used',
  'summary.noConditions': 'no conditions',
  'summary.askValue': '?',
  'summary.updated': 'changed {date}',

  'plural.column': ['column', 'columns'],
  'plural.condition': ['condition', 'conditions'],
  'plural.row': ['row', 'rows'],
  'plural.range': ['range', 'ranges'],

  // Webview
  'ui.savedFilter': 'Saved filter',
  'ui.noFilter': '— none —',
  'ui.matchingFilters': 'Matching this file',
  'ui.otherFilters': 'Other filters',
  'ui.manageFilters': 'Manage saved filters…',
  'ui.modified': 'modified',
  'ui.modifiedHint': 'Your changes apply to this export only until you save them in the filter.',
  'ui.saveChanges': 'Save changes',
  'ui.saveAsFilter': 'Save as filter…',
  'ui.saveAsNew': 'Save as new…',
  'ui.revert': 'Revert',
  'ui.copy': 'Copy',
  'ui.openInEditor': 'Open in editor',
  'ui.export': 'Export .md',

  'ui.section.source': 'Data',
  'ui.section.columns': 'Columns',
  'ui.section.conditions': 'Row conditions',
  'ui.section.options': 'Options',

  'ui.sheet': 'Sheet',
  'ui.hiddenSheet': '{name} (hidden)',
  'ui.selection': 'Selection',
  'ui.selectionPlaceholder': 'e.g. A4:G120',
  'ui.selectionHint': 'Drag over the cells. Ctrl/⌘ + drag adds another range, Shift + click extends, double-click selects the whole block of data.',
  'ui.noSelection': 'Nothing selected yet. Select cells in the sheet, or pick a saved filter.',
  'ui.excelTables': 'Excel tables',
  'ui.usedRange': 'All data on the sheet',
  'ui.toEnd': 'to the end of the data',
  'ui.toEndHint': 'Rows added later are included when the filter is used again.',
  'ui.removeRange': 'Remove range',
  'ui.headerRow': 'First row of the selection holds the column names',
  'ui.multiRange': 'Several ranges',
  'ui.multiRange.auto': 'one table when the columns match',
  'ui.multiRange.merge': 'always one table',
  'ui.multiRange.separate': 'a table per range',

  'ui.columnsAll': 'All',
  'ui.columnsNone': 'None',
  'ui.columnsHint': 'Drag to change the order.',
  'ui.columnsMissing': 'Not in this file: {names}',
  'ui.columnsNew': 'New in this file (not selected): {names}',
  'ui.noColumns': 'Select data to see its columns.',
  'ui.sortBy': 'Sort by',
  'ui.sortNone': 'as in the sheet',
  'ui.sortAsc': 'ascending',
  'ui.sortDesc': 'descending',

  'ui.match': 'Keep rows that meet',
  'ui.match.all': 'all conditions',
  'ui.match.any': 'any condition',
  'ui.addCondition': 'Add condition',
  'ui.removeCondition': 'Remove condition',
  'ui.noConditions': 'No conditions — every row is kept.',
  'ui.valueFixed': 'Fixed',
  'ui.valueAsk': 'Ask when used',
  'ui.valueFixedHint': 'The value is saved with the filter.',
  'ui.valueAskHint': 'The value is chosen each time the filter is used.',
  'ui.askPlaceholder': 'chosen when the filter is used',
  'ui.valuePlaceholder': 'value',
  'ui.numberPlaceholder': 'number',
  'ui.datePlaceholder': 'date, e.g. 2026-03-15',
  'ui.choose': 'choose…',
  'ui.any': 'any',
  'ui.search': 'Search',
  'ui.addValue': 'Use “{value}”',
  'ui.selectedCount': '{count} selected',
  'ui.clear': 'Clear',
  'ui.conditionMissing': 'Column “{name}” is not in this file; the condition is skipped.',

  'op.eq': '=',
  'op.neq': '≠',
  'op.contains': 'contains',
  'op.notContains': 'does not contain',
  'op.startsWith': 'starts with',
  'op.endsWith': 'ends with',
  'op.gt': '>',
  'op.gte': '≥',
  'op.lt': '<',
  'op.lte': '≤',
  'op.empty': 'is empty',
  'op.notEmpty': 'is not empty',

  'ui.opt.skipHiddenRows': 'Skip hidden rows',
  'ui.opt.skipHiddenColumns': 'Skip hidden columns',
  'ui.opt.skipEmptyRows': 'Skip empty rows',
  'ui.opt.dropEmptyColumns': 'Leave out columns with no values',
  'ui.opt.fillMerged': 'Repeat the value of merged cells',
  'ui.opt.values': 'Values',
  'ui.opt.values.formatted': 'as shown in Excel',
  'ui.opt.values.raw': 'without number formats',
  'ui.opt.layout': 'Layout',
  'ui.opt.layout.table': 'table',
  'ui.opt.layout.sections': 'section per row',
  'ui.opt.title': 'Heading',
  'ui.opt.title.none': 'none',
  'ui.opt.title.sheet': 'sheet name',
  'ui.opt.title.filter': 'filter name',
  'ui.opt.title.custom': 'own text',
  'ui.opt.titleText': 'Heading text',
  'ui.opt.lineBreaks': 'Line breaks in cells',
  'ui.opt.lineBreaks.br': '<br>',
  'ui.opt.lineBreaks.space': 'space',
  'ui.opt.alignNumbers': 'Align numbers to the right',
  'ui.opt.padColumns': 'Line up the table in the source',

  'ui.preview': 'Markdown preview',
  'ui.previewCode': 'Markdown',
  'ui.previewRendered': 'Rendered',
  'ui.previewEmpty': 'The Markdown appears here once data is selected.',
  'ui.previewNoRows': 'No row meets the conditions.',
  'ui.previewTruncated': 'Showing the first {count} rows; the export has all of them.',
  'ui.stats': '{kept} of {total} {rows}',
  'ui.statsColumns': '{count} {columns}',

  'ui.apply': 'Apply',
  'ui.cancel': 'Cancel',
  'ui.applyTitle': 'Filter “{name}”',
  'ui.applyHint': 'Choose the values for this use. Leave a field empty to keep every value.',
  'ui.note.sheetByHeaders': 'Sheet “{sheet}” does not have the data; found it on “{used}”.',
  'ui.note.sheetMissing': 'There is no sheet “{sheet}”; using “{used}”.',
  'ui.note.headersPartly': 'Found {found} of {total} column names at {range}.',
  'ui.note.headersNotFound': 'The column names were not found; using the saved range {ref}.',
  'ui.loading': 'Reading the workbook…',
  'ui.cellCount': '{rows} × {cols}',

  'column.letter': 'Column {letter}',
};

const PL = {
  'action.overwrite': 'Nadpisz',
  'action.delete': 'Usuń',
  'action.rename': 'Zmień nazwę',
  'action.openSettings': 'Otwórz ustawienia',

  'dialog.openLabel': 'Otwórz',
  'dialog.filter': 'Skoroszyty Excela',
  'dialog.saveLabel': 'Eksportuj',
  'dialog.markdown': 'Markdown',

  'error.notXlsx': '{file} nie jest skoroszytem Excela (.xlsx).',
  'error.oldXls': '{file} jest w starym formacie .xls. Otwórz go w Excelu, zapisz jako .xlsx (Skoroszyt programu Excel) i spróbuj ponownie.',
  'error.failed': 'Nie udało się odczytać {file}: {reason}',
  'error.noSheets': '{file} nie ma arkuszy z komórkami.',

  'info.exported': 'Zapisano {file}.',
  'info.copied': 'Skopiowano Markdown do schowka.',
  'info.filterSaved': 'Zapisano filtr „{name}”.',
  'info.filterUpdated': 'Zaktualizowano filtr „{name}”.',
  'info.filterDeleted': 'Usunięto filtr „{name}”.',
  'info.noFilters': 'Nie ma jeszcze zapisanych filtrów. Otwórz plik Excela, wybierz potrzebne dane i zapisz to ustawienie jako filtr.',

  'prompt.filterName': 'Nazwa filtra',
  'prompt.filterNamePlaceholder': 'np. Specyfikacja komunikatów',
  'prompt.filterNameEmpty': 'Podaj nazwę filtra.',
  'prompt.filterExists': 'Filtr „{name}” już istnieje. Zastąpić go?',
  'prompt.replace': 'Zastąp',
  'prompt.deleteFilter': 'Usunąć filtr „{name}”? Tego nie da się cofnąć.',
  'prompt.renameFilter': 'Nowa nazwa filtra',
  'prompt.nameTaken': 'Filtr „{name}” już istnieje.',

  'manage.title': 'Zapisane filtry',
  'manage.placeholder': 'Zapisane filtry są dostępne w każdym projekcie. Przyciskami zmienisz nazwę albo usuniesz filtr.',

  'summary.sheet': 'arkusz {sheet}',
  'summary.columns': '{count} {columns}',
  'summary.conditions': '{count} {conditions}',
  'summary.ask': '{count} podawane przy użyciu',
  'summary.noConditions': 'bez warunków',
  'summary.askValue': '?',
  'summary.updated': 'zmieniony {date}',

  'plural.column': ['kolumna', 'kolumny', 'kolumn'],
  'plural.condition': ['warunek', 'warunki', 'warunków'],
  'plural.row': ['wiersz', 'wiersze', 'wierszy'],
  'plural.range': ['zakres', 'zakresy', 'zakresów'],

  'ui.savedFilter': 'Zapisany filtr',
  'ui.noFilter': '— brak —',
  'ui.matchingFilters': 'Pasujące do tego pliku',
  'ui.otherFilters': 'Pozostałe filtry',
  'ui.manageFilters': 'Zarządzaj zapisanymi filtrami…',
  'ui.modified': 'zmieniony',
  'ui.modifiedHint': 'Zmiany obowiązują tylko w tym eksporcie, dopóki nie zapiszesz ich w filtrze.',
  'ui.saveChanges': 'Zapisz zmiany',
  'ui.saveAsFilter': 'Zapisz jako filtr…',
  'ui.saveAsNew': 'Zapisz jako nowy…',
  'ui.revert': 'Przywróć',
  'ui.copy': 'Kopiuj',
  'ui.openInEditor': 'Otwórz w edytorze',
  'ui.export': 'Eksportuj .md',

  'ui.section.source': 'Dane',
  'ui.section.columns': 'Kolumny',
  'ui.section.conditions': 'Warunki dla wierszy',
  'ui.section.options': 'Opcje',

  'ui.sheet': 'Arkusz',
  'ui.hiddenSheet': '{name} (ukryty)',
  'ui.selection': 'Zaznaczenie',
  'ui.selectionPlaceholder': 'np. A4:G120',
  'ui.selectionHint': 'Przeciągnij po komórkach. Ctrl/⌘ + przeciągnięcie dodaje kolejny zakres, Shift + klik rozszerza, dwuklik zaznacza cały blok danych.',
  'ui.noSelection': 'Nic jeszcze nie zaznaczono. Zaznacz komórki w arkuszu albo wybierz zapisany filtr.',
  'ui.excelTables': 'Tabele Excela',
  'ui.usedRange': 'Wszystkie dane arkusza',
  'ui.toEnd': 'do końca danych',
  'ui.toEndHint': 'Wiersze dopisane później zostaną uwzględnione przy ponownym użyciu filtra.',
  'ui.removeRange': 'Usuń zakres',
  'ui.headerRow': 'Pierwszy wiersz zaznaczenia zawiera nazwy kolumn',
  'ui.multiRange': 'Kilka zakresów',
  'ui.multiRange.auto': 'jedna tabela, gdy kolumny się zgadzają',
  'ui.multiRange.merge': 'zawsze jedna tabela',
  'ui.multiRange.separate': 'osobna tabela dla każdego zakresu',

  'ui.columnsAll': 'Wszystkie',
  'ui.columnsNone': 'Żadna',
  'ui.columnsHint': 'Przeciągnij, aby zmienić kolejność.',
  'ui.columnsMissing': 'Brak w tym pliku: {names}',
  'ui.columnsNew': 'Nowe w tym pliku (niezaznaczone): {names}',
  'ui.noColumns': 'Zaznacz dane, aby zobaczyć ich kolumny.',
  'ui.sortBy': 'Sortuj według',
  'ui.sortNone': 'jak w arkuszu',
  'ui.sortAsc': 'rosnąco',
  'ui.sortDesc': 'malejąco',

  'ui.match': 'Zostaw wiersze, które spełniają',
  'ui.match.all': 'wszystkie warunki',
  'ui.match.any': 'dowolny warunek',
  'ui.addCondition': 'Dodaj warunek',
  'ui.removeCondition': 'Usuń warunek',
  'ui.noConditions': 'Brak warunków — zostają wszystkie wiersze.',
  'ui.valueFixed': 'Stała',
  'ui.valueAsk': 'Podawana przy użyciu',
  'ui.valueFixedHint': 'Wartość zostanie zapisana w filtrze.',
  'ui.valueAskHint': 'Wartość wybierasz za każdym razem, gdy używasz filtra.',
  'ui.askPlaceholder': 'wybierana przy użyciu filtra',
  'ui.valuePlaceholder': 'wartość',
  'ui.numberPlaceholder': 'liczba',
  'ui.datePlaceholder': 'data, np. 2026-03-15',
  'ui.choose': 'wybierz…',
  'ui.any': 'dowolna',
  'ui.search': 'Szukaj',
  'ui.addValue': 'Użyj „{value}”',
  'ui.selectedCount': 'wybrane: {count}',
  'ui.clear': 'Wyczyść',
  'ui.conditionMissing': 'Kolumny „{name}” nie ma w tym pliku; warunek jest pomijany.',

  'op.eq': '=',
  'op.neq': '≠',
  'op.contains': 'zawiera',
  'op.notContains': 'nie zawiera',
  'op.startsWith': 'zaczyna się od',
  'op.endsWith': 'kończy się na',
  'op.gt': '>',
  'op.gte': '≥',
  'op.lt': '<',
  'op.lte': '≤',
  'op.empty': 'jest puste',
  'op.notEmpty': 'nie jest puste',

  'ui.opt.skipHiddenRows': 'Pomijaj ukryte wiersze',
  'ui.opt.skipHiddenColumns': 'Pomijaj ukryte kolumny',
  'ui.opt.skipEmptyRows': 'Pomijaj puste wiersze',
  'ui.opt.dropEmptyColumns': 'Pomijaj kolumny bez wartości',
  'ui.opt.fillMerged': 'Powtarzaj wartość scalonych komórek',
  'ui.opt.values': 'Wartości',
  'ui.opt.values.formatted': 'tak jak w Excelu',
  'ui.opt.values.raw': 'bez formatów liczb',
  'ui.opt.layout': 'Układ',
  'ui.opt.layout.table': 'tabela',
  'ui.opt.layout.sections': 'sekcja dla każdego wiersza',
  'ui.opt.title': 'Nagłówek',
  'ui.opt.title.none': 'brak',
  'ui.opt.title.sheet': 'nazwa arkusza',
  'ui.opt.title.filter': 'nazwa filtra',
  'ui.opt.title.custom': 'własny tekst',
  'ui.opt.titleText': 'Tekst nagłówka',
  'ui.opt.lineBreaks': 'Nowe linie w komórkach',
  'ui.opt.lineBreaks.br': '<br>',
  'ui.opt.lineBreaks.space': 'spacja',
  'ui.opt.alignNumbers': 'Wyrównuj liczby do prawej',
  'ui.opt.padColumns': 'Wyrównuj kolumny tabeli w kodzie',

  'ui.preview': 'Podgląd Markdown',
  'ui.previewCode': 'Markdown',
  'ui.previewRendered': 'Widok',
  'ui.previewEmpty': 'Markdown pojawi się tutaj, gdy zaznaczysz dane.',
  'ui.previewNoRows': 'Żaden wiersz nie spełnia warunków.',
  'ui.previewTruncated': 'Widać pierwsze {count} wierszy; eksport zawiera wszystkie.',
  'ui.stats': '{kept} z {total} {rows}',
  'ui.statsColumns': '{count} {columns}',

  'ui.apply': 'Zastosuj',
  'ui.cancel': 'Anuluj',
  'ui.applyTitle': 'Filtr „{name}”',
  'ui.applyHint': 'Wybierz wartości na ten raz. Puste pole oznacza dowolną wartość.',
  'ui.note.sheetByHeaders': 'Arkusz „{sheet}” nie zawiera tych danych; znaleziono je w „{used}”.',
  'ui.note.sheetMissing': 'Nie ma arkusza „{sheet}”; użyto „{used}”.',
  'ui.note.headersPartly': 'Znaleziono {found} z {total} nazw kolumn w {range}.',
  'ui.note.headersNotFound': 'Nie znaleziono nazw kolumn; użyto zapisanego zakresu {ref}.',
  'ui.loading': 'Czytam skoroszyt…',
  'ui.cellCount': '{rows} × {cols}',

  'column.letter': 'Kolumna {letter}',
};

const TRANSLATIONS = { pl: PL };

// 'pl-PL' → 'pl'; anything unknown falls back to the English base.
function baseLanguage(tag) {
  return String(tag || 'en').toLowerCase().split(/[-_]/)[0];
}

const PLURAL_RULES = {
  en: (n) => (n === 1 ? 0 : 1),
  pl: (n) => {
    if (n === 1) return 0;
    const d = n % 10;
    const h = n % 100;
    return d >= 2 && d <= 4 && (h < 12 || h > 14) ? 1 : 2;
  },
};

// The full message table for a language, English filled in where needed —
// what the webview receives.
function messagesFor(tag) {
  const table = TRANSLATIONS[baseLanguage(tag)] || {};
  const out = {};
  for (const key of Object.keys(EN)) {
    const v = table[key];
    out[key] = (typeof v === 'string' && v !== '') || (Array.isArray(v) && v.length) ? v : EN[key];
  }
  return out;
}

function fromMessages(messages, lang) {
  const rule = PLURAL_RULES[lang] || PLURAL_RULES.en;
  const t = (key, params) => {
    const text = messages[key];
    if (typeof text !== 'string') return key;
    return text.replace(/\{(\w+)\}/g, (whole, name) =>
      params && name in params ? String(params[name]) : whole);
  };
  const plural = (key, n) => {
    const forms = messages[key];
    if (!Array.isArray(forms) || !forms.length) return typeof forms === 'string' ? forms : key;
    return forms[Math.min(rule(n), forms.length - 1)];
  };
  return { lang: PLURAL_RULES[lang] ? lang : 'en', t, plural };
}

function forLanguage(tag) {
  const lang = baseLanguage(tag);
  return fromMessages(messagesFor(tag), PLURAL_RULES[lang] ? lang : 'en');
}

module.exports = { forLanguage, fromMessages, messagesFor, baseLanguage, EN, TRANSLATIONS };
