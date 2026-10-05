'use strict';

const { check, eq } = require('./harness');
const F = require('../src/core/filter');

const cell = (text, kind = 's', num = null) => ({ text, kind, num });
const cond = (op) => ({ op });

check('equality is case- and space-insensitive', () => {
  eq(F.test(cond('eq'), cell(' crm '), ['CRM']), true);
  eq(F.test(cond('eq'), cell('CRM'), ['FUND', 'crm']), true);
  eq(F.test(cond('eq'), cell('CORE'), ['FUND', 'crm']), false);
  eq(F.test(cond('neq'), cell('DELETED'), ['deleted']), false);
  eq(F.test(cond('neq'), cell('OK'), ['deleted']), true);
});

check('booleans in any spelling', () => {
  eq(F.test(cond('eq'), cell('TRUE', 'b', 1), ['true']), true);
  eq(F.test(cond('eq'), cell('Tak'), ['TRUE']), true);
  eq(F.test(cond('eq'), cell('no'), ['FALSE']), true);
  eq(F.test(cond('eq'), cell('FALSE', 'b', 0), ['TRUE']), false);
});

check('numbers compare as numbers', () => {
  eq(F.test(cond('eq'), cell('1,234.50', 'n', 1234.5), ['1234.5']), true);
  eq(F.test(cond('gt'), cell('10', 'n', 10), ['9']), true);
  eq(F.test(cond('lt'), cell('10', 'n', 10), ['9']), false);
  eq(F.test(cond('gte'), cell('1 234,5'), ['1234,5']), true);
  eq(F.test(cond('lte'), cell('abc'), ['5']), false);
});

check('dates compare as dates', () => {
  const d = cell('2023-03-15', 'd', 45000.5);
  eq(F.test(cond('eq'), d, ['2023-03-15']), true);
  eq(F.test(cond('eq'), d, ['15.03.2023']), true);
  eq(F.test(cond('gt'), d, ['2023-03-14']), true);
  eq(F.test(cond('lt'), d, ['2023-03-15']), false);
  eq(F.test(cond('gte'), d, ['2023-03-15']), true);
});

check('text operators', () => {
  eq(F.test(cond('contains'), cell('Customer id'), ['ID']), true);
  eq(F.test(cond('notContains'), cell('Customer id'), ['name', 'ID']), false);
  eq(F.test(cond('startsWith'), cell('PRJ-12'), ['prj']), true);
  eq(F.test(cond('endsWith'), cell('PRJ-12'), ['13']), false);
});

check('empty and not empty', () => {
  eq(F.test(cond('empty'), cell(''), []), true);
  eq(F.test(cond('empty'), null, []), true);
  eq(F.test(cond('notEmpty'), cell('x'), []), true);
});

check('a condition without a value does nothing', () => {
  eq(F.isActive(cond('eq'), []), false);
  eq(F.isActive(cond('eq'), ['  ']), false);
  eq(F.isActive(cond('empty'), []), true);
  eq(F.isActive(cond('contains'), ['a']), true);
});

check('column types and their operators', () => {
  eq(F.columnType([cell('1', 'n', 1), cell('2', 'n', 2), cell('')]), 'number');
  eq(F.columnType([cell('2023-01-01', 'd', 44927)]), 'date');
  eq(F.columnType([cell('TRUE', 'b', 1), cell('FALSE', 'b', 0)]), 'bool');
  eq(F.columnType([cell('a'), cell('1', 'n', 1)]), 'text');
  eq(F.operatorsFor('text').includes('gt'), false);
  eq(F.operatorsFor('number').includes('gt'), true);
  eq(F.operatorsFor('bool').includes('contains'), false);
});

check('parsing numbers', () => {
  eq(F.parseNumber('1,234.5'), 1234.5);
  eq(F.parseNumber('1.234,5'), 1234.5);
  eq(F.parseNumber('12,5'), 12.5);
  eq(F.parseNumber('50%'), 0.5);
  eq(F.parseNumber('1 000'), 1000);
  eq(F.parseNumber('12a'), null);
  eq(F.parseNumber(''), null);
});

require('./harness').done('filter');
