
const assert = require('assert');
const Sheet = require('./sheet.js');



const sheet = new Sheet();

sheet.values = [
  [new Date(2021,2,1), "date1"],
  [new Date(2021,2,2), "date2"],
  [,,],
];


const test_values = [
  [new Date(2021,2,1), "date1"],
  [new Date(2021,2,2), "date2"],
];


assert.deepStrictEqual(sheet.getRange(1,1,2,2).getValues(), test_values);
assert.deepStrictEqual(sheet.getRange('A1:B1').getValues(), [[new Date(2021,2,1), "date1"]]);
assert.deepStrictEqual(sheet.getRange('A1:B').getValues(),
   [
    [new Date(2021,2,1), "date1"],
    [new Date(2021,2,2), "date2"],
    [undefined,undefined,],
  ]
);
  
