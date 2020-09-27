#!/usr/bin/env node

const assert = require('assert');
const rewire = require('rewire');
const Code = rewire('../../report_generator/Code.js');

const libraryGet = Code.__get__('libraryGet');

const tests = new Map();

tests.set('dummy_func',() => {
  const dummy_func = libraryGet('dummy_func');
  assert.equal(dummy_func(10), 20, 'dummy_func failed');
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

  const fieldObject = validator.getField(validArgs[0]);

  // field object expected properties
  ['name', 'index', 'type', 'minValue', 'maxValue', 'exactValues']
    .forEach((prop, i) => {
      assert.strictEqual(fieldObject.hasOwnProperty(prop), true, `property ${prop} not found`);
      assert.strictEqual(fieldObject[prop], validArgs[i]);
  });
  return;

  // test field name
  assert.strictEqual(validator.getField(validArgs[0]).name, validArgs[0]);
  // test field index
  assert.strictEqual(validator.getField(validArgs[0]).index, validArgs[1]);
  return;
  // test field type
  assert.strictEqual(validator.getField(validArgs[0]).name, validArgs[0]);
  // test field min value
  assert.strictEqual(validator.getField(validArgs[0]).name, validArgs[0]);
  // test field max value
  assert.strictEqual(validator.getField(validArgs[0]).name, validArgs[0]);
  // test field exact values
  assert.strictEqual(validator.getField(validArgs[0]).name, validArgs[0]);
  
  // test readonly
  assert.throws(()=>{validator.getField(validArgs[0]).name = 'changed';},
    {name: 'Error', reason: 'readonly'});
});



tests.get('dummy_func')();
tests.get('FieldValidator')();
