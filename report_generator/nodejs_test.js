#!/usr/bin/env node

const assert = require('assert');
const rewire = require('rewire');
const Code = rewire('./Code.js');

const libraryGet = Code.__get__('libraryGet');

const tests = new Map();

tests.set('dummy_func',() => {
  const dummy_func = libraryGet('dummy_func');
  assert.equal(dummy_func(10), 20, 'dummy_func failed');
});




tests.get('dummy_func')();
