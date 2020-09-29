#!/usr/bin/env node

const assert = require('assert');
const rewire = require('rewire');
const Code = rewire('../../report_generator/Code.js');


const libraryGet = Code.__get__('libraryGet');

// usefull
const getType = libraryGet('getType');

const tests = new Map();

tests.set('getType', () => {
  const getType = libraryGet('getType');

  // primitive types
  const _number = 1;
  const _string = 'some string literal';
  const _boolean = true;
  const _null = null;
  const _undefined = undefined;
  // structural types
  const _function = function(){};
  const _array = [];
  const _object = {};
  const _map = new Map();
  const _set = new Set();
  assert.strictEqual(getType(_number), 'number');
  assert.strictEqual(getType(_string), 'string');
  assert.strictEqual(getType(_boolean), 'boolean');
  assert.strictEqual(getType(_null), 'null');
  assert.strictEqual(getType(_undefined), 'undefined');
  assert.strictEqual(getType(_function), 'Function');
  assert.strictEqual(getType(_array), 'Array');
  assert.strictEqual(getType(_object), 'Object');
  assert.strictEqual(getType(_map), 'Map');
  assert.strictEqual(getType(_set), 'Set');
});

tests.set('argumentsValidator', () => {
  const argumentsValidator = libraryGet('argumentsValidator');
  const [setArgTypes, validateArgs] = argumentsValidator();

  assert.throws(()=>{setArgTypes(1,'string')},
    {name:'TypeError'}, "setArgTypes can be called only with strings");

  function InvalidType(){};
  const someValidArgs = ['str',1,true,{},[],new Map(),new Set(),new Date()];

  // proper definition of function using argumentsValidator
  function funcDef(s,n,b,O,A,M,S,D){
    const [setArgTypes, validateArgs] = argumentsValidator();
    setArgTypes('string','number','boolean','Object','Array','Map','Set','Date');
    validateArgs(...[...arguments]);
    const body = "rest of the body of the function definition";
  }
  // test for normal valid arguments
  assert.doesNotThrow(()=>{funcDef(...someValidArgs);});
  // test for wrong number of arguments
  assert.throws(()=>{
    const fewerArgs = someValidArgs.slice(0, -1);
    const moreArgs = [...someValidArgs,new Date()];
    funcDef(...fewerArgs);
    funcDef(...moreArgs);
    },{name:'TypeError',});
  // test for wrong type of one argument
  someValidArgs.map((arg, i) =>{
    const currArgType = typeof arg==='object'?arg.constructor.name:typeof arg;
    const oneInvalidType = [...someValidArgs];
    oneInvalidType[i] = new InvalidType();
      assert.throws(()=>{
        funcDef(...oneInvalidType);
        },{name:'TypeError', expectedType: currArgType});
    });

});

tests.set('FieldValidator', () => {
  const FieldValidator = libraryGet('FieldValidator');

  const exactIntValues = new Set();
  const exactStrValues = new Set();
  exactIntValues.add(3); exactIntValues.add(4);
  exactStrValues.add('c'); exactStrValues.add('d');
  class CustomType{};

  const testArguments = {
    setField:{
      correctArgs: [
        ['nums',0,'number'],
        ['nums',0,'number',0],
        ['nums',0,'number',null],
        ['nums',0,'number',null,2],
        ['nums',0,'number',0,null],
        ['nums',0,'number',0,2],
        ['nums',0,'number',null,null],
        ['nums',0,'number',null,null,exactIntValues]
      ],
      wrongNumOfArgs: [
        [],
        ['nums',0],
        ['nums',0,'number',null,null,exactIntValues,'anotherOne']
      ],
      wrongTypeOfArgs: [
        ['nums',0,'number','str'],
        ['nums',0,'number',0,'str'],
        ['nums',0,'number',null,'str'],
        ['nums',0,'number','str',null],
        ['nums',0,'number','str','str'],
        ['nums',0,'number','str','str',exactIntValues],
        ['nums',0,'number',null,'str',exactIntValues],
        ['nums',0,'number','str',null,exactIntValues],
        ['nums',0,'number',null,null,exactStrValues],
      ],
    },
    validate:{
      correctArgs: [
        ['ints', 3],
        ['objs', {}],
        ['customs', new CustomType()],
      ],
      wrongNumOfArgs: [
        ['ints'],
        ['objs', {}, 'one more'],
      ],
      wrongTypeOfArgs: [
        [3,2],
        [{},2],
      ],
    },
  };

  // test correct arguments
  for (const args of testArguments.setField.correctArgs){
    assert.doesNotThrow(()=>{
      new FieldValidator().setField(...args); 
    })
  }
  // test wrong number of arguments
  for (const args of testArguments.setField.wrongNumOfArgs){
    assert.throws(()=>{
      new FieldValidator().setField(...args);
    },{name:'TypeError'});
  }
  // test wront types of arguments
  for (const args of testArguments.setField.wrongTypeOfArgs){
    assert.throws(()=>{
      new FieldValidator().setField(...args);
    },{name:'TypeError'});
  }
  // test correct arguments
  for (const args of testArguments.validate.correctArgs){
    assert.doesNotThrow(()=>{
      new FieldValidator()
        .setField(args[0],1,getType(args[1]))
        .validate(...args); 
    })
  }
  // test wrong number of arguments
  for (const args of testArguments.validate.wrongNumOfArgs){
    assert.throws(()=>{
      new FieldValidator()
        .setField(args[0],1,getType(args[1]))
        .validate(...args);
    },{name:'TypeError'});
  }
  // test wront types of arguments
  for (const args of testArguments.validate.wrongTypeOfArgs){
    assert.throws(()=>{
      new FieldValidator()
        .setField('fn',0,'number')
        .validate(...args);
    },{name:'TypeError'});
  }
  // test same name field
  assert.throws(()=>{
    new FieldValidator().setField('name',0,'number').setField('name',1,'number');
  },{name:'Error',message:'fieldName already exists'});
  // test form same index
  assert.throws(()=>{
    new FieldValidator().setField('name1',0,'number').setField('name2',0,'number');
  },{name:'Error',message:'fieldIndex already exists'});

  const exactDateValues = new Set();
  const date1 = new Date(), date2 = new Date(), date3 = new Date();
  date2.setDate(date1.getDate()+1); date3.setDate(date2.getDate()+1);
  exactDateValues.add(date2); exactDateValues.add(date3);
  const allTypesExcept_undefined = [null,true,'str',[],{},new Map(),new Set(),new Date()];
  const allTypesExcept_null = [undefined,true,'str',[],{},new Map(),new Set(),new Date()];
  const allTypesExcept_bool = [undefined,null,'str',[],{},new Map(),new Set(),new Date()];
  const allTypesExcept_str = [undefined,null,true,1,[],{},new Map(),new Set(),new Date()];
  const allTypesExcept_int = [undefined,null,true,'str',[],{},new Map(),new Set(),new Date()];
  const allTypesExcept_date = [undefined,null,true,'str',[],{},new Map(),new Set()];
  // test validate for return ant throw with different inputs
  // for these expects to return true
  const validator = new FieldValidator()
  let caseNum = -1;
  for (const [fieldDescription, correctValues, wrongValues, wrongValueTypes] of [
    [['num',0,'number'],[-1,0,1,1.2,10000],[],allTypesExcept_int],
    [['num',0,'number',0],[0,1,2,10000],[-1,-10000],allTypesExcept_int],
    [['num',0,'number',null,2],[-10000,-1,0,1,2],[3,4],allTypesExcept_int],
    [['num',0,'number',0,2],[0,1,2],[-1,3],allTypesExcept_int],
    [['num',0,'number',0,null],[0,1,2,10000],[-1,-2],allTypesExcept_int],
    [['num',0,'number',null,null,exactIntValues],[3,4],[-1,0,2,5],allTypesExcept_int],
    [['date',0,'Date'],[date1,date2,date3,new Date()],[],allTypesExcept_date],
    [['date',0,'Date',date2],[date2,date3],[date1],allTypesExcept_date],
    [['date',0,'Date',date1,date2],[date1,date2],[date3],allTypesExcept_date],
    [['date',0,'Date',null,date2],[date1,date2],[date3],allTypesExcept_date],
    [['date',0,'Date',date2,null],[date2,date3],[date1],allTypesExcept_date],
    [['date',0,'Date',null,null,exactDateValues],[date2,date3],[date1],allTypesExcept_date],
    [['str',0,'string'],['str','a','b','c'],[],allTypesExcept_str],
    [['str',0,'string',null,null],['str','a','b','c'],[],allTypesExcept_str],
    [['str',0,'string','c'],['c','d'],['a','b'],allTypesExcept_str],
    [['str',0,'string','b','d'],['b','c','d'],['a','e'],allTypesExcept_str],
    [['str',0,'string',null,'d'],['b','c','d'],['e','f'],allTypesExcept_str],
    [['str',0,'string','c',null],['c','d','e'],['a','b'],allTypesExcept_str],
    [['str',0,'string',null,null,exactStrValues],['c','d'],['a','b','e'],allTypesExcept_str],
  ]){
    ++caseNum;
    fieldDescription[0] += caseNum;
    fieldDescription[1] = caseNum;
    validator.setField(...fieldDescription);
    const fieldName = fieldDescription[0];
    correctValues.map(testValue => {
      assert.strictEqual(validator.validate(fieldName, testValue),true);
    });
    wrongValues.map(testValue =>{
      assert.throws(()=>{validator.validate(fieldName, testValue)},{name:'ValueError'},
      `field description ${fieldDescription}, wrong test value is ${testValue}`);
    });
    wrongValueTypes.map(testValue =>{
      assert.throws(()=>{validator.validate(fieldName, testValue)},{name:'TypeError'},
      `field description ${fieldDescription}, wrong test value type is ${testValue}`);
    });
  }
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
tests.get('getType')();

