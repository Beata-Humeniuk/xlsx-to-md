'use strict';

const { check, eq } = require('./harness');
const { formatNumber, isDateFormat } = require('../src/core/numfmt');

const cases = [
  [1234.5, 'General', '1234.5'],
  [0.1 + 0.2, null, '0.3'],
  [1234.5, '#,##0.00', '1,234.50'],
  [1234.5, '0', '1235'],
  [-5, '0.00', '-5.00'],
  [-12, '#,##0;(#,##0)', '(12)'],
  [0, '0;-0;"zero"', 'zero'],
  [0.256, '0.0%', '25.6%'],
  [0.5, '0%', '50%'],
  [1234567, '0.00E+00', '1.23E+06'],
  [0.000123, '0.0E+00', '1.2E-04'],
  [12345, '#,##0,"k"', '12k'],
  [3, '000', '003'],
  [1234, '00-00', '12-34'],
  [5, '"Qty: "0', 'Qty: 5'],
  [42.1, '[$€-407] #,##0.00', '€ 42.10'],
  [42, '#,##0.00 [$PLN]', '42.00 PLN'],
  [1.5, '# ?/?', '1 1/2'],
  [0.75, '?/4', '3/4'],
  [7, '[Red]0;[Blue]-0', '7'],
  [150, '[>100]"big";"small"', 'big'],
  [5, '[>100]"big";"small"', 'small'],
  [12.3, '0.0_);(0.0)', '12.3 '],
  [2.5, '#.##', '2.5'],
  [2, '#.00', '2.00'],
];

for (const [value, code, expected] of cases) {
  check('number ' + value + ' as ' + code, () => eq(formatNumber(value, code, false), expected));
}

const dates = [
  [45000, 'yyyy-mm-dd', '2023-03-15'],
  [45000.5, 'd.mm.yyyy hh:mm', '15.03.2023 12:00'],
  [45000, 'dd mmm yyyy', '15 Mar 2023'],
  [45000, 'dddd, mmmm d', 'Wednesday, March 15'],
  [0.5, 'h:mm AM/PM', '12:00 PM'],
  [0.75, 'hh:mm:ss', '18:00:00'],
  [1.25, '[h]:mm', '30:00'],
  [0.000694444, 'mm:ss', '01:00'],
  [1, 'yyyy-mm-dd', '1900-01-01'],
  [60, 'yyyy-mm-dd', '1900-02-29'],
  [61, 'yyyy-mm-dd', '1900-03-01'],
  [45000.0000116, 'hh:mm:ss.0', '00:00:01.0'],
];

for (const [value, code, expected] of dates) {
  check('date ' + value + ' as ' + code, () => eq(formatNumber(value, code, false), expected));
}

check('1904 date system', () => eq(formatNumber(0, 'yyyy-mm-dd', true), '1904-01-01'));

check('date formats are recognized', () => {
  eq(isDateFormat('yyyy-mm-dd'), true);
  eq(isDateFormat('h:mm'), true);
  eq(isDateFormat('[h]:mm:ss'), true);
  eq(isDateFormat('0.00'), false);
  eq(isDateFormat('"m"0'), false);
  eq(isDateFormat('0.00E+00'), false);
  eq(isDateFormat('General'), false);
});

require('./harness').done('numfmt');
