'use strict';

// A specification-like workbook shared by the extraction and resolution tests.

const { xlsx } = require('./xlsx-builder');
const { openWorkbook } = require('../src/workbook');

const HEAD = ['Field', 'Type', 'Required', 'Description', 'Status', 'Owner', 'System', 'Direction', 'Amount'];

function specRows() {
  return [
    ['Message specification'],
    [],
    HEAD,
    ['id', 'string', true, 'Identifier', 'OK', 'anna', 'CRM', 'IN', 10],
    ['name', 'string', false, 'Display name\nshown in lists', 'OK', 'ewa', 'CRM', 'OUT', 2.5],
    ['old_code', 'int', true, 'Legacy', 'DELETED', 'anna', 'CRM', 'IN', 7],
    ['amount', 'decimal', true, 'Value | net', 'OK', 'piotr', 'FUND', 'IN', 1200],
    ['items', 'List<Item>', true, '*Required* list', 'OK', 'ewa', 'CRM', 'IN', 3],
    [],
    ['Notes below the table'],
  ];
}

function specBook(extra = {}) {
  return xlsx([Object.assign({ name: 'Messages', rows: specRows(), merges: ['A1:D1'] }, extra)]);
}

function loadSpec(extra) {
  return openWorkbook(specBook(extra));
}

module.exports = { HEAD, specRows, specBook, loadSpec };
