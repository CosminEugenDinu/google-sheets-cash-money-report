#!/usr/bin/env node

const assert = require('assert');
const rewire = require('rewire');
const Code = rewire('../../report_generator/Code.js');


const libraryGet = Code.__get__('libraryGet');

const addMessages = libraryGet('addMessages');
const Log = libraryGet('Log');
const getType = libraryGet('getType');
const argumentsValidator = libraryGet('argumentsValidator');
const FieldValidator = libraryGet('FieldValidator');
const validateRecord = libraryGet('validateRecord');
const cleanRawData = libraryGet('cleanRawData');
const renderReport = libraryGet('renderReport');

const getRecords = libraryGet('getRecords');

const tests = new Map();

tests.set('addMessages', () => {
  function someProcedure(...args){
    const addMessage = addMessages(someProcedure);
    addMessage(`I was called with first arg: ${args[0]}`);
    // do some work
    addMessage('I\'m done');
    return;
  }

  someProcedure('arg1');

  expectedMessages = new Map();
  expectedMessages.set(`I was called with first arg: arg1`,[1]);
  expectedMessages.set('I\'m done',[1]);

  assert.deepStrictEqual(someProcedure.messages, expectedMessages); 
});

tests.set('Log', () => {
  // mocks
  const cellValue = []
  const cell = {
    setValue(msg){cellValue[0] = msg;},
    setBackground(c){},
    setFontColor(c){},
    setVerticalAlignment(a){}};
  const range = {
    clear(){return this;},
    merge(){return this;},
    getCell(x,y){return cell;},};
  const sheet = {
    getRange(x,y,z,w){return range;}};

  const location = [1,1,5,3]; 
  let log = Log(sheet, location);
  let messages = ['ceva', 'altceva'];
  log(messages[0]); log(messages[1]);
  assert.equal(cellValue[0], '> '+messages.join('\n> '));

  cellValue[0] = undefined;
  log = Log(sheet, location);
  messages = ['cva'];
  log(messages[0]);
  assert.equal(cellValue[0], '> '+messages.join('\n> '));
});

tests.set('getType', () => {
  // primitive types
  const _number = 1;
  const _string = 'some string literal';
  const _boolean = true;
  const _null = null;
  const _undefined = undefined;
  const _NaN = NaN;
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
  assert.strictEqual(getType(_NaN), 'nan');
  assert.strictEqual(getType(_function), 'Function');
  assert.strictEqual(getType(_array), 'Array');
  assert.strictEqual(getType(_object), 'Object');
  assert.strictEqual(getType(_map), 'Map');
  assert.strictEqual(getType(_set), 'Set');
});

tests.set('argumentsValidator', () => {
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
  const exactIntValues = new Set();
  const exactStrValues = new Set();
  exactIntValues.add(3); exactIntValues.add(4);
  exactStrValues.add('c'); exactStrValues.add('d');
  class CustomType{};

  const testArguments = {
    setField:{
      correctArgs: [
        ['nums0',0,'number'],
        ['nums1',0,'number',undefined],
        ['nums',0,'number',undefined,undefined],
        ['nums',0,'number',undefined,undefined,undefined],
        ['nums2',0,'number',undefined,2],
        ['nums',0,'number',2,undefined],
        ['nums',0,'number',undefined,null],
        ['nums',0,'number',null,undefined],
        ['nums',0,'number',null],
        ['nums',0,'number',null,2],
        ['nums',0,'number',0,null],
        ['nums',0,'number',0,2],
        ['nums',0,'number',0,2,null],
        ['nums',0,'number',null,null],
        ['nums',0,'number',null,null,null],
        ['nums',0,'number',null,null,exactIntValues],
        ['nums',0,'number',undefined,undefined,exactIntValues],
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
        ['nums',0,'number',1,3,exactIntValues],
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
      try { new FieldValidator().setField(...args); }
      catch (e) {
        console.log(JSON.stringify(e));
        throw e;
      }
    }, `should not throw if args are ${args}`)
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
  let y = 2020, m = 9, d = 1;
  const date1=new Date(y,m,d), date2=new Date(y,m,++d), date3=new Date(y,m,++d);
  const invalidDate = new Date('xxx');
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
    [['date',0,'Date'],[date1,date2,date3,new Date()],[invalidDate],allTypesExcept_date],
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
      assert.throws(()=>{validator.validate(fieldName, testValue)},
        {name:'ValueError'},
      `field description ${fieldDescription}, wrong test value is ${testValue}`);
    });
    wrongValueTypes.map(testValue =>{
      assert.throws(()=>{validator.validate(fieldName, testValue)},
        {name:'TypeError'},
      `field description ${fieldDescription}, wrong test value type is ${testValue}`);
    });
  }
});

tests.set('validateRecord', () => {
  const record = [1,2,3,4,5,6,7];
  const fields = {date:0,name:1};
});

tests.set('cleanRawData', () => {
  // set verbosity to 0...2
  const v = 0;
  cleanRawData.verbosity = v;

  for (const [correctable_case, correct_case] of [
    [
    // correct record
    [new Date(2015,1,20), 'ref0', 'docType0', 'descr0', 0, 20],
    [new Date(2015,1,20), 'ref0', 'docType0', 'descr0', 0, 20]
    ],
    [
    // wrong ref type
    [new Date(2015,1,21), 21, 'docType1', 'docDesrc1', 1, 21],
    [new Date(2015,1,21), '21', 'docType1', 'docDesrc1', 1, 21],
    ],
    [
    // wrong doc_type type
    [new Date(2015,1,22),'ref2', 222, 'docDescr2', 0, 22],
    [new Date(2015,1,22),'ref2', '222', 'docDescr2', 0, 22],
    ],
    [
    // wrong descr type
    [new Date(2015,1,23),'ref3', 'docType3', 333, 0, 23],
    [new Date(2015,1,23),'ref3', 'docType3', '333', 0, 23],
    ],
    [
    // wrong I_O_type type
    [new Date(2015,1,24), 'ref4', 'docType4', 'descr4', '0', 24],
    [new Date(2015,1,24), 'ref4', 'docType4', 'descr4', 0, 24],
    ],
    [
    // wrong value type
    [new Date(2015,1,25), 'ref5', 'docType5', 'descr5', 0, '555'],
    [new Date(2015,1,25), 'ref5', 'docType5', 'descr5', 0, 555],
    ],
    [
    // wrong date type
    ['2015,1,26', 'ref6', 'docType6', 'descr6', 0, 26],
    [new Date('2015,1,26'), 'ref6', 'docType6', 'descr6', 0, 26],
    ],
  ]){
    // mock
    const fromDate = new Date(), toDate = new Date();
    const company = new Map();
    let values = [
      ['date','ref','doc_type','descr','I_O_type','value'],
      [...correctable_case],
    ];
    const Range = {
      getValues(){return values;},
      setValues(newValues){values = newValues;}};
    const rawDataSheet = {
      getRange(str){return Range;},getName(){return 'sheet_name'}};

    cleanRawData(fromDate, toDate, company, rawDataSheet);
    v && console.log(cleanRawData.messages);

    // now values should be cleaned
    assert.deepStrictEqual(values[1], correct_case);
  }

  for (const throwingCase of [
    // wrong I_O_type value
    [new Date(2015,1,27), 'ref7', 'docType7', 'descr7', 3, 26],
    // wrong I_O_type value and type
    [new Date(2015,1,28), 'ref8', 'docType8', 'descr8', '8', 26],
    // wrong date value and type
    ['xx,yy,zz', 'ref9', 'docType9', 'descr9', 1, 26],
    // wrong value type and value
    [new Date(2015,1,28), 'ref8', 'docType8', 'descr8', '8', 'xx'],
  ]){
    // mocks
    const fromDate = new Date(), toDate = new Date();
    const company = new Map();
    let values = [
      ['date','ref','doc_type','descr','I_O_type','value'],
      [...throwingCase],
    ];
    const Range = {
      getValues(){return values;},
      setValues(newValues){values = newValues;}};
    const rawDataSheet = {
      getRange(str){return Range;},getName(){return 'sheet_name'}};

    assert.throws(()=>{
      cleanRawData(fromDate, toDate, company, rawDataSheet);
    },{name:'ValueError'},`should throw ValueError on case ${throwingCase}`);
    v && console.log(cleanRawData.messages);
  }
  // 10 empty rows (if encountered) is considered end of data set
  const twelveEmptyRows = Array(12).fill(Array(6));
  const fiveEmptyRows = Array(5).fill(Array(6));
  for (const [unsortedDuplicates, sortedWithoutDuplicates] of [
    [
      [
      ['date','ref','doc_type','descr','I_O_type','value'],
      [new Date(2015,1,28), 'ref8', 'docType8', 'descr8', 0, 26],
      [new Date(2015,1,28), 'ref8', 'docType8', 'descr8', 0, 26],
      ...fiveEmptyRows,
      [new Date(2015,1,28), 'ref8', 'docType8', 'descr7', 0, 26],
      [new Date(2015,1,27), 'ref8', 'docType8', 'descr8', 0, 26],
      ...fiveEmptyRows,
      [new Date(2015,1,27), 'ref8', 'docType8', 'descr8', 0, 26],
      ...twelveEmptyRows,
      ],
      [
      ['date','ref','doc_type','descr','I_O_type','value'],
      [new Date(2015,1,27), 'ref8', 'docType8', 'descr8', 0, 26],
      [new Date(2015,1,28), 'ref8', 'docType8', 'descr8', 0, 26],
      [new Date(2015,1,28), 'ref8', 'docType8', 'descr7', 0, 26],
      [ , , , , , ,],
      [ , , , , , ,],
      ...fiveEmptyRows,
      ...fiveEmptyRows,
      ...twelveEmptyRows,
      ]
    ],
  ]) {
    // mocks
    const fromDate = new Date(), toDate = new Date();
    const company = new Map();
    let values = [...unsortedDuplicates];

    const Range = {
      getValues(){return values;},
      setValues(newValues){values = newValues;}};
    const rawDataSheet = {
      getRange(str){return Range;},getName(){return 'sheet_name'}};

    cleanRawData.verbosity = v; 
    cleanRawData(fromDate, toDate, company, rawDataSheet);

    assert.deepStrictEqual(values, sortedWithoutDuplicates);
    v && console.log(cleanRawData.messages);
  }
});

tests.set('renderReport', () => {
  renderReport.verbosity = 2;
  //renderReport();
  //console.log(renderReport.messages);
});

tests.set('getRecords', () => {
  // verbosity
  const v = 2;

  const fieldDescriptors = [
    {fieldName:"date",fieldType:"Date"},
    {fieldName:"ref",fieldType:"string"},
    {fieldName:"doc_type",fieldType:"string"},
    {fieldName:"descr",fieldType:"string"},
    {fieldName:"I_O_type",fieldType:"number",minValue:0,maxValue:1},
    {fieldName:"value",fieldType:"number"}
  ];

  const twelveEmptyRows = Array(12).fill(Array(12));
  const fiveEmptyRows = Array(5).fill(Array(12));
  let values = [
    ['ref',,,'date','doc_type','descr','strange_field','I_O_type',,'value',,,],
    ['ref8',,,new Date(2015,1,28),'docType8','descr8','unknown',0,,26,,,],
    ['ref8',,,new Date(2015,1,28),'docType8','descr8','unknown',0,,26,,,],
    ...fiveEmptyRows,
    ['ref8',,,new Date(2015,1,28),'docType8','descr7','unknown',0,,26,,,],
    ['ref8',,,new Date(2016,1,27),'docType8','descr8','unknown',0,,26,,,],
    ['ref8',,,new Date(2016,1,27),'docType8','descr8','unknown',0,,26,,,],
    ...twelveEmptyRows,
    ];
  // mocks
  const Range = {getValues(){return values;}};
  const rawDataSheet = {getRange(str){return Range;},getName(){return 'sheetName';}};
  
  const fieldNames = {
    'ref':0,'date':3,'doc_type':4,'descr':5,'I_O_type':7,'value':9
  };
  
  // expected return from procedure getRecords
  const records = new Map();
  // for every unique or duplicate date.toJSON set and array of {Map} records

  let row_i = 0;
  while(++row_i < values.length){
    const row = values[row_i];
    // exclude empty row (assuming first undefined value)
    if (row[0]===undefined) continue;
    const date = row[3];
    const dateKey = date.toJSON();

    const record = new Map(); // {Map} record
    for (const fieldName in fieldNames){
      const fieldIndex = fieldNames[fieldName];
      record.set(fieldName,fieldName==='date'?dateKey:row[fieldIndex]);
    }
    if (records.has(dateKey)) records.get(dateKey).push(record)
    else records.set(dateKey, [record]); 
  }
  getRecords.verbosity = v;
  assert.deepStrictEqual(getRecords(rawDataSheet, fieldDescriptors), records);
  console.log(getRecords.messages);
});

tests.get('addMessages')();
tests.get('Log')();
tests.get('getType')();
tests.get('argumentsValidator')();
tests.get('FieldValidator')();
tests.get('validateRecord')();
tests.get('cleanRawData')();
tests.get('renderReport')();
tests.get('getRecords')();
