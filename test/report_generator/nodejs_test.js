#!/usr/bin/env node

const assert = require('assert');
const rewire = require('rewire');
const Code = rewire('../../report_generator/Code.js');


const libraryGet = Code.__get__('libraryGet');

const tests = new Map();

tests.set('argumentsValidator', () => {
  const argumentsValidator = libraryGet('argumentsValidator');
  const [setArgTypes, validateArgs] = argumentsValidator();

  assert.throws(()=>{setArgTypes(1,'string')},
    {name:'TypeError'}, "setArgTypes can be called only with strings");

  function InvalidType(){};
  const someValidTypes = ['str',1,true,{},[],new Map(),new Set(),new Date()];

  // proper definition of function using argumentsValidator
  function properDef(s,n,b,O,A,M,S,D){
    const [setArgTypes, validateArgs] = argumentsValidator();
    setArgTypes('string','number','boolean','Object','Array','Map','Set','Date');
    validateArgs(...[...arguments]);
    
    const body = "rest of the body of the function definition";
  }
  
  // wrong definition of function using argumentsValidator
  // number of formal parameters not the same as num of arguments of setArgTypes
  function wrongDef1(s,n,b,O,A,M,S) 


  assert.doesNotThrow(()=>{
    properDef();});
  assert.throws(()=>{
    },{name:'TypeError',});
  assert.throws(()=>{
    properDef(1,2,true);
  },{name:'TypeError',});

});

tests.set('FieldValidator', () => {
  const FieldValidator = libraryGet('FieldValidator');
  const validator = new FieldValidator();
  const validArgs = [
    'nums', // fieldName
    0, // fieldIndex
    'number', // fieldType
    -1, // minValue
    10, // maxValue
    [9, 10, 12].reduce((set,exactVal)=>{set.add(exactVal); return set;}, new Set())
  ];

  validator.setField(...validArgs);

  const field = validator.getFieldByName(validArgs[0]);

  // field object expected properties
  ['name', 'index', 'type', 'minValue', 'maxValue', 'exactValues']
    .forEach((prop, i) => {
      assert.strictEqual(field.has(prop), true, `property-key ${prop} not found`);
      assert.strictEqual(field.get(prop), validArgs[i]);
  });
});

tests.set('Log', () => {
  // mocks
  const cellValue = []
  const cell = {setValue(msg){cellValue[0] = msg;}, setBackground(c){}, setFontColor(c){}, setVerticalAlignment(a){}};
  const range = {clear(){return this;}, merge(){return this;}, getCell(x,y){return cell;},};
  const sheet = {getRange(x,y,z,w){return range;}};
  
  const Log = libraryGet('Log');

  let cellPos = [1,1], defaultVerbosity = 3;
  let [v, log] = Log(sheet, cellPos, defaultVerbosity);
  let messages = ['ceva', 'altceva'];
  v(1); log(messages[0]); log(messages[1]);
  assert.equal(cellValue[0], '> '+messages.join('\n> '));

  cellValue[0] = undefined;
  [v, log] = Log(sheet, cellPos, 2);
  messages = ['cva'];
  v(2); log(messages[0]);
  assert.equal(cellValue[0], '> '+messages.join('\n> '));

  cellValue[0] = undefined;
  [v, log] = Log(sheet, cellPos, 1);
  messages = ['cva'];
  v(1); log(messages[0]);
  assert.equal(cellValue[0], '> '+messages.join('\n> '));

  cellValue[0] = undefined;
  [v, log] = Log(sheet, cellPos, 1);
  messages = ['cva'];
  v(2); log(messages[0]);
  // because 2 from 'v(2)' is >= 1 from defaultVerbosity
  assert.equal(cellValue[0], undefined);
  v(1); log('important');
  // this is logged because it is important! kidding...
  assert.equal(cellValue[0], '> important');
});


tests.get('FieldValidator')();
tests.get('Log')();
tests.get('argumentsValidator')();

