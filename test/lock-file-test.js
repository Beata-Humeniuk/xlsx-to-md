'use strict';

const { check, eq } = require('./harness');
const { isLockFile, workbookFor } = require('../src/lockFile');

check('lock files are recognized', () => {
  eq(isLockFile('~$Book1.xlsx'), true);
  eq(isLockFile('Book1.xlsx'), false);
  eq(isLockFile('a~$b.xlsx'), false);
});

check('the workbook with the same name', () => {
  eq(workbookFor('~$Book1.xlsx', ['Book1.xlsx', '~$Book1.xlsx', 'Other.xlsx']), 'Book1.xlsx');
});

check('the workbook whose first two characters were replaced', () => {
  eq(workbookFor('~$ex ante.xlsx', ['Plex ante.xlsx', '~$ex ante.xlsx']), 'Plex ante.xlsx');
  eq(workbookFor('~$P1694_Zestawienie_Wymagan.xlsm', ['P1694_Zestawienie_Wymagan.xlsm']), 'P1694_Zestawienie_Wymagan.xlsm');
});

check('unknown or ambiguous', () => {
  eq(workbookFor('~$ex ante.xlsx', ['Notes.xlsx']), null);
  eq(workbookFor('~$ex ante.xlsx', ['Abex ante.xlsx', 'Cdex ante.xlsx']), null);
  eq(workbookFor('~$ab.xlsx', ['xxab.xlsx', 'yyab.xlsx']), null);
});

require('./harness').done('lock-file');
