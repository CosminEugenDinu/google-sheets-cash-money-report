
/**
 * Main function, call from sheet button, where it is assigned
 * This function is considered global
 */

function main () {

const Settings = libraryGet('settings');
const getRecords = libraryGet('getRecords');
const Log = libraryGet('Log');

// global settings variables
const REPORT_GENERATOR_SPREADSHEET_ID = "1nDNPcpgP9TlAcTSKASwFxuQQswdAI7Wm3I8Z-7rSGnE";
const INTERFACE_SHEET_NAME = "Interface";
const SETTINGS_SHEET_NAME = "settings";
const RAWDATA_SHEET_SUFFIX = "_rawdata";


  
const repGenSprSheet = SpreadsheetApp.openById(REPORT_GENERATOR_SPREADSHEET_ID);

// interfaceSheet is a {Sheet} with selects and button that triggers 'makeReport'
const interfaceSheet = repGenSprSheet.getSheetByName(INTERFACE_SHEET_NAME);
const settingsSheet = repGenSprSheet.getSheetByName(SETTINGS_SHEET_NAME);

const verbosity = interfaceSheet.getSheetValues(8,5,1,1)[0][0];
// v=0 - critical, v=1 - informal, v=2 - too verbose 
// function to check some level against verbosity (which is set on interface)
const v = level => verbosity === level; 
// instantiate log function
const log = Log(interfaceSheet, [10,5,8,3]);

const settings = new Settings(settingsSheet, 50, 1000);

const varIndex = settings
    .getField('procedure.variable.name')
    .getByValue('renderReport.reportspreadsheetid');
const REPORT_SPREADSHEET_ID = settings
    .getField('procedure.variable.value')
    .getByIndex(varIndex);


const companies = getCompanies(settingsSheet);
const companyAliases = Array.from(companies.keys());
const computedRawDataSheetNames = companyAliases.map(
  alias => alias + RAWDATA_SHEET_SUFFIX); 
const rawDataSheets = repGenSprSheet.getSheets().filter(
  sheet => {
  const sheetName = sheet.getSheetName();
  // get all sheets except Interface and settingsSheet
  return (sheetName !== INTERFACE_SHEET_NAME) && (sheetName !== SETTINGS_SHEET_NAME);
});
  
// if company alias name was changed in settingsSheet
updateRawDataSheetNames(rawDataSheets, computedRawDataSheetNames);

// user selects in interfaceSheet
const procedure = interfaceSheet.getSheetValues(8,1,1,1)[0][0];
const companyAlias = interfaceSheet.getSheetValues(8,2,1,1)[0][0];
const [fromDate, toDate] = interfaceSheet.getSheetValues(8,3,1,2)[0];
const company = companies.get(companyAlias);

// data source sheet corresponds with chosen company alias from drop-down
const rawDataSheet = repGenSprSheet.getSheetByName(companyAlias+RAWDATA_SHEET_SUFFIX);

//---------------------------------------------------------------------------------

if (procedure==='renderReport'){
  let args;
  try{
    getRecords.verbosity = verbosity;
    const dataRecords = getRecords(rawDataSheet);
    const messages = getRecords.messages;
    if (messages) for (let [msg, count] of messages){ log(count, msg); }

    const template = defaultTemplate();
    const targetSpreadsheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);

    args = [fromDate, toDate, company, dataRecords, template, targetSpreadsheet];

  } catch(e){
    throw new Error(
      `When initializing arguments for procedure ${procedure}, got:\n`+
      `e.message: ${e.message}, json: ${JSON.stringify(e)}`);
  }
  let mesages;
  try{
    const renderReport = libraryGet(procedure);
    renderReport.verbosity = verbosity;
    renderReport(...args);
    messages = renderReport.messages;
  } catch (e) {
    throw new Error(`Procedure ${procedure} failed with:\n${e.message}`+
    `\nComplete Error object is:\n${JSON.stringify(e)}`);
  }
  if (messages)
    for (const [msg, count] of messages) log(count, msg);
}

//---------------------------------------------------------------------------------

if (procedure==='importData'){
  let args;
  try{
    // spreadsheet links iterable
    const dataLinks = settings.getField(`link.${companyAlias}`).getValues();
    // this pattern is uset to search records in source spreadsheets (dataLinks);
    const identifierPattern = settings.getField('procedure.variable.value')
      .getByIndex(
        settings.getField('procedure.variable.name')
        .getByValue('importData.identifierPattern')
      );
    const [sheetToImportTo] = rawDataSheets.filter(
      sheet => sheet.getName() === company.get('alias')+RAWDATA_SHEET_SUFFIX
      );

    args = [fromDate, toDate, company, dataLinks, identifierPattern, sheetToImportTo];

  } catch(e){
    throw new Error(
      `When initializing arguments for procedure ${procedure}, got:\n`+
      `e.message: ${e.message}, json: ${JSON.stringify(e)}`);
  }

  let mesages;
  try{
    const importData = libraryGet(procedure);
    importData.verbosity = verbosity;
    importData(...args);
    messages = importData.messages;
  } catch (e) {
    throw new Error(`Procedure ${procedure} failed with:\n${e.message}`+
    `\nComplete Error object is:\n${JSON.stringify(e)}`);
  }
  if (messages)
    for (const [msg, count] of messages) log(count, msg);
}

//---------------------------------------------------------------------------------

if (procedure === 'cleanRawData'){

  const args = [fromDate, toDate, company, rawDataSheet];

  try{
    const cleanRawData = libraryGet(procedure);
    cleanRawData.verbosity = verbosity;
    cleanRawData(...args);
    const messages = cleanRawData.messages;
    if (messages) for (const [msg, count] of messages){log(count, msg);}
  } catch (e) {
    throw new Error(`Procedure ${procedure} failed with:\n${e.message}`+
    `\nComplete Error object is:\n${JSON.stringify(e)}`);
  }
}

} // main END
//---------------------------------------------------------------------------------


// -------------------------- library --------------------------------


// ----------Global functions (in makeReport scope)---------------

/**
 * Runs in O(n^2), unfortunately
 * @param {Array} records_1 
 * @param {Array} records_2
 * @returns {Array} merged
 */
function mergeDateRecords(records_1, records_2){
  const merged = [];
  records_1.map(map1 => {
    let alsoInRecords_2 = false;
    for (const map2 of records_2)
      if (areTheSame(map1, map2)){
         //log(`Record ${Array.from(map1.values())} is duplicate, so skipped`); 
        alsoInRecords_2 = true;
        break;
      }
    if (! alsoInRecords_2)
      merged.push(map1);
  }
  );
  records_2.map(
    map2 => merged.push(map2)
  );
  return merged;
}

/**
 * Compares two {Map} instances
 */
function areTheSame(map_1, map_2){
  for (const [k, v] of map_1.entries())
    if (v !== map_2.get(k)){
      return false;
    }
  return true;
}

/**
 * @param {Spreadsheet} spreadsheet
 * @param {Number} rowLim - maximum number of rows to search
 * @param {Number} colLim - maximum number of columns to search
 * @returns {Map} records
 *      - {string} keys - dates (ISO 8601)
 *      - {Array} values - of {Map} records, like {'date'=>{Date}, 'ref'=>32, etc.} 
 */
function searchRecords(spreadsheet, identifierPattern, rowLim=50, colLim=6){
  const records = new Map();

  // measurements
  const messages = new Map();

  // pattern to search against 
  //const identifierRe = /=RIGHT\(CELL\("filename",A\d\),LEN\(CELL\("filename",A\d\)\)-FIND\("\]",CELL\("filename",A\d\)\)\)/;
  const identifierRe = new RegExp(identifierPattern);

  for (const sheet of spreadsheet.getSheets()){
    
    const sheetRecords = [];
    
    // looks in first column for identifier (which is a formula)
    const searchRange = sheet.getRange(1,1,rowLim,1);
    // {Array[][]} formulas
    const formulas = searchRange.getFormulas();

    // iterate over first column and search for pattern
    let row_i = -1;
    while(++row_i < formulas.length){
      // if pattern is found, then look 5 columns right for record
      if (formulas[row_i][0].match(identifierRe)){
        // if record has at least one value, then is a valid record 
        const [record] = sheet.getSheetValues(row_i+1, 2, 1, colLim-1);
        isValidRecord(record) && sheetRecords.push(record);
      }
    }
    

    // if some records were found, add them to {Map} records 
    if (sheetRecords.length){
      // assume sheetName is a date string like '01.02.2020'; 
      const [d, m, y] = sheet.getName().split('.');
      const dateStr = new Date(+y, +m-1, +d).toJSON();
      records.set(dateStr, []);

      sheetRecords.map(
        record => {
          const recordMap = new Map();
          const [ref, doc_type, descr, input, output] = record;

          recordMap.set('date', dateStr);
          recordMap.set('ref', ref || null);
          recordMap.set('doc_type', doc_type || null);
          recordMap.set('descr', descr || null);
          if (input)
            recordMap.set('I_O_type', 1);
          else if (output)
            recordMap.set('I_O_type', 0);
          recordMap.set('value', input || output || null);

          records.get(dateStr).push(recordMap);
        }
      );
     
    } else {
      const message = 'Records not found';
      if (messages.has(message))
        messages.get(message).push({sheet: sheet.getName(), spreadsheet: spreadsheet.getName()})  
      else
        messages.set(message, []);
    }
  }

// log accumulated messages
//messages.forEach((vals, mess) => v(1) && log(mess, vals.length, JSON.stringify(vals)));
return records;
}

/**
 * Verifies if a record is valid
 * 
 * @param {Array[]} record - contains three values
 * @returns {Boolean} isValid 
 */
function isValidRecord(record){

  const len = 5;
  if (record.length !== len) return false;

  // if at least one value is truthy, then is a valid record
  let i = -1;
  while (++i < len)
    if (record[i]) return true;

  return false;
}

/**
 * @param {string} link - google sheet url link
 * @returns {string} id - spreadsheet id
 */
function extractId(link){
  const urlReStr = 'https\:\/\/docs.google.com\/spreadsheets.*\/d\/';
  const idReStr = '(?<id>.*)';
  const restReStr = '\/.*';
  const sheetIdRe = new RegExp(`${urlReStr}${idReStr}${restReStr}`);
  const {id} = link.match(sheetIdRe).groups;
  return id;
}


/**
 * @param {Date} date1
 * @param {Date} date2
 * @returns {Array} of {string} dates - all dates (ISO 8601) between date1 and date2
 */
function datesBetween(date1, date2){
  const dates = [];
  const newDate = new Date(date1.getTime());

  if (date1.getTime() <= date2.getTime()){
    dates.push(newDate.toJSON());
    while (newDate.getTime() < date2.getTime()){
      newDate.setDate(newDate.getDate() + 1);
      dates.push(newDate.toJSON());
    }
  } else {
    dates.push(newDate.toJSON());
    while (newDate.getTime() > date2.getTime()){
      newDate.setDate(newDate.getDate() - 1);
      dates.push(newDate.toJSON());
    }    
  }

  return dates;
}

function getCompanies(sheet, records=10, fields=4){
  const companies = new Map();
  let company_id = 0
  for (const row of sheet.getSheetValues(2,1,records,fields)){
    if (row.filter(val => val === "").length){
    company_id ++;
    continue;
    }
  const company = new Map();
  const alias = row[0];
  company.set('alias', alias);
  company.set('name', row[1]);
  company.set('tax_id', row[2]);
  company.set('reg_num', row[3]);
  companies.set(alias, company)
  company_id ++;
 }
return companies;
}

function updateRawDataSheetNames(rawDataSheets, computedNames){
      rawDataSheets.map(
        (sheet, i) => sheet.setName(computedNames[i])
      )
}



function isObject(type){
  return typeof(type) === 'object' && type.constructor.name === 'Object';
}

/**
 * Replaces '{}' from string with value
 * 
 * @param {string} templateString
 * @param {*} value - will replace '{}'
 * @returns {string} - replaced string
 */
function replaceCurly(templateString, value){
  return templateString.replace(/\{\s*\}/, value.toString());
}

/**
 * Returns default template for generating report
 */
function defaultTemplate(){

  const LABEL_STYLE = {fontSize:8, horizontalAlignment:'center', verticalAlignment:'middle', background:'lightgray', borders:[null, null, null, null, false, false]};
  const TARGET_STYLE = {borders:[null, null, null, null, false, false]};

  const TEMPLATE = {
    _layoutRange:[1,1,50,6],
    _columnWidths:[75,80,30,250,80,80],
    _rowHeight:15,
    companyName:{
      label_element:{cell:[2,2], value:"Societatea",
        style:{fontSize:8, fontWeight:"bold", horizontalAlignment:'left'}},
      target_element:{cell:[2,4]}},
    tax_id:{
      label_element:{cell:[3,2], value:"CUI",
        style:{fontSize:8, fontWeight:'bold', horizontalAlignment:'left'}},
      target_element:{cell:[3,4],
        style:{horizontalAlignment:'left'}}},
    reg_num:{
      label_element:{cell:[4,2], value:"Nr. Reg. Com.:", 
        style:{fontSize:8, fontWeight:'bold', horizontalAlignment:'left'}},
      target_element:{cell:[4,4]}},
    title:{
      label_element:{cell:[8,2], value:"REGISTRUL DE CASA", offset:[2,4],
        style:{fontSize:12, fontWeight:'bold', horizontalAlignment:'center', verticalAlignment:'middle'}}},
    document:{
      label_element:{cell:[11,1], value:"Document", offset:[1, 3],
        style:LABEL_STYLE}},
    explanations:{
      label_element:{cell:[11,4], value:"EXPLICATII", offset:[2,1],
        style:LABEL_STYLE}},
    input:{
      label_element:{cell:[11,5], value:"INCASARI", offset:[2,1],
        style:LABEL_STYLE}},
    output:{
      label_element:{cell:[11,6], value:"PLATI", offset:[2,1],
        style:LABEL_STYLE}},
    date:{
      label_element:{cell:[12,1], value:"DATA",
        style:LABEL_STYLE}},
    ref:{
      label_element:{cell:[12,2], value: "NR",
        style:LABEL_STYLE}},
    doc_type:{
      label_element:{cell:[12,3], value:"TIP",
      style:LABEL_STYLE}},
    previous_balance:{
      label_element:{cell:[13,4], value:"SOLD LUNA/ZIUA PRECEDENTA"},
      target_element:{cell:[13,5]}},
    record:{
      date:{
        target_element:{cell:[14,1]},
      },
      ref:{
        target_element:{cell:[14,2]},
      },
      doc_type:{
        target_element:{cell:[14,3]},
      },
      descr:{
        target_element:{cell:[14,4]},
      },
      input:{
        target_element:{cell:[14,5]},
      },
      output:{
        target_element:{cell:[14,6],
        style: TARGET_STYLE},
      }
    },
    day_total_label:{
      label_element:{cell:[15,4],value:"Total la data de {}:",
        style:LABEL_STYLE}
    },
    day_total_targets:{
      total_input:{
        target_element:{cell:[15,5]}
      },
      total_day_output:{
        target_element:{cell:[15,6]}
      }
    },
    day_balance:{
      label_element:{cell:[16,4],value:"Sold la data de {}:",
        style:LABEL_STYLE},
      target_element:{cell:[16,5]}
    },
    body:{
      frame_element:{cell:[13,1], extent:[4,6],
        style:{borders:[null, true, true, true, false, false]}},
    },
  };
  return TEMPLATE;
}


function libraryGet(required){

// library members dictionary
const library = new Map();

/**
 * Takes function object, adds own property 'messages',
 * and returns closure (used to add debugging messages).
 *
 * @param {function} procedure
 *
 * @return {function} addMessage
 * @prop {Map} procedure.messages
 */
function addMessages(procedure){
  if (typeof procedure !== 'function')
    throw new TypeError(`${procedure.name} is not function`);

  procedure.messages = new Map();

  const addMessage = message =>
    procedure.messages.has(message) ?
      ++procedure.messages.get(message)[0] :
      procedure.messages.set(message,[1]);

  return addMessage;
}

/**
 * Logging function - logs to specified cell
 *      - instantiate with const log = Log(sheet, [x,y,xOffset,yOffset);
 *      - usage: log("Welcome to log console!");
 * @param {Sheet} sheet
 * @param {number[]} location
 */
function Log(sheet, location){

  // clear console space
  //sheet.getRange(...location).clear();
  //const range = sheet.getRange(...cellPos,8,3).merge();
  const range = sheet.getRange(...location).clear().merge();
  const cell = range.getCell(1,1);
  cell.setBackground("black");
  cell.setFontColor("white");
  cell.setVerticalAlignment("top");
  
  const _logs = [];

  const log = (...args) => {
    _logs.push(args.toString());
    cell.setValue('> '+_logs.join('\n> '));
    // returns true to permit chaining like: log('something') && another_statement;
    return true;
  }
  

  //return [setVerbosity, log];
  return log;
}


class Settings{
  constructor(settingsSheet, rowLim, colLim){
    const range = settingsSheet.getRange(1,1,rowLim,colLim);
    const sheetValues = settingsSheet.getSheetValues(1,1,rowLim,colLim);

    // key {string} fieldName, value {Number} index of column
    const fieldNames = new Map();
    sheetValues[0].forEach(
      (fieldName, index) => {
        if (!!fieldName) fieldNames.set(fieldName, index); 
      }
    );
    
    const fields = new Map();
    for (const [fieldName, index] of fieldNames){
      const values = new Map();
      const indexes = new Map();

      for (let i=1; i<sheetValues.length; i++){
        const row = sheetValues[i];
        const fieldValue = row[index];
        if ( ! fieldValue) continue;
        values.set(i, fieldValue);
        indexes.set(fieldValue, i);
      }
      
      const fieldObject = {
        getByIndex(index){
          if ( ! values.has(index))
            throw new ReferenceError(`Field ${fieldName} does not have key-index ${index}`);
          return values.get(index);
        },
        getByValue(val){
          if ( ! indexes.has(val))
            throw new ReferenceError(`Field ${fieldName} does not have key-value ${val}`);
          return indexes.get(val);
        },
        getValues(){
          return values.values();
        }
      }

      //fields.set(fieldName, values);
      fields.set(fieldName, fieldObject);
    }

    this._sheetValues = sheetValues;
    this._range = range;
    this._fieldNames = fieldNames;
    this._fields = fields;

  }


  get fieldNames(){
    return Array.from(this._fieldNames.keys());
  }
  set fieldNames(val){
    throw new Error(`Settings.fieldNames is read only. Cannot set ${val}.`)
  }

  getField(fieldName){
    if ( ! this._fieldNames.has(fieldName))
      throw new ReferenceError(`${fieldName} is not a field name`);
    return this._fields.get(fieldName);
  }
} // class Settings END


function renderReport(
  fromDate, toDate, company, dataRecords, template,targetSpreadsheet){

  const addMessages = libraryGet('addMessages');

  // initialize debug messaging
  const thisProcedure = renderReport;
  const addMessage = addMessages(thisProcedure); // adds prop {Map} messages
  // @prop {number} verbosity could be added at procedure call
  const v = thisProcedure.verbosity ? thisProcedure.verbosity : 0;
  
  v>0 && addMessage('Procedure renderReport begin');

  /**
   * Class Element - is a piece of sheet... (cell, range)
   *
   * Depending on type of typeKey (e.g. 'target_element', 'label_element', 'frame_element'),
   * assigns specific properties (e.g. only element 'frame_element' has property 'extent')
   * When render method is called, that element produces effect on target sheet,
   * like setting a value in a cell or changing background color.
   */
  class Element {

    /**
     * @param {Object} elem - a TEMPLATE object containing properties specific to type=typeKey 
     * @param {string} typeKey - key in {Map} tree where elem is stored
     *   - represents the type of element;
     *   - supported types can be verified with Element.getSupportedTypes();
     */
    constructor(typeKey, elem){

      const typesProps = Element._typesProps

      if (!this.supportedTypes.includes(typeKey))
        throw new TypeError(`${typeKey} is not a valid element type`+
          `supported types are: ${this.supportedTypes}`);
      this._type = typeKey;

      // reference to template element object
      this._templateElement = elem;

      for (const prop of Element._typesProps.get(typeKey))
        this[prop] = null;
     
      for (const p in elem){
        if (typesProps.get(this._type).includes(p)){
          // if property is a reference to an array, so copy it
          if (Array.isArray(elem[p])) this[p] = [...elem[p]];
          else this[p] = elem[p];
        } else {
          throw new TypeError(`Property ${p} is not a supported by element type ${this._type}`);
          }
      }

    }

    get templateElement(){
      return this._templateElement;
    }

    get type(){
      return this._type;
    }

    get supportedTypes(){
      return Element._supportedTypes;
    }

    static getSupportedTypes(){
      return Element._supportedTypes;
    }

    /**
     * Converts tuple array like [1, 2] into {string} key '1:2'
     */
    get keyCell(){
      const [x, y] = this.cell;
      const key = `${x}:${y}`;
      return key;
    }

    /**
     * Sets properties on range objects (e.g. set borders on cells in sheet)
     *
     * @param {Range} range - instance returned by sheet.getRange(x, y, ...)
     * @param {string} property - key in {Map} properties (a local variable)
     * @param {string|Number|Array} value - required by a range method
     */
    static setProperty(range, property, value){
      const properties = new Map();
      properties.set("background", ()=>range.setBackground(value));
      properties.set("borders", ()=>range.setBorder(...value));
      properties.set("fontSize", ()=>range.setFontSize(value));
      properties.set("fontColor", ()=>range.setFontColor(value));
      properties.set("fontWeight", ()=>range.setFontWeight(value));
      properties.set("horizontalAlignment", ()=>range.setHorizontalAlignment(value));
      properties.set("verticalAlignment", ()=>range.setVerticalAlignment(value));
      
      // set property on range object
      properties.get(property)();
    }

    render(sheet){
      if (!this.cell) throw new TypeError(`Cannot render element when element.cell=${cell}`);

      let range = sheet.getRange(...this.cell);
      //range.clear();
      if (this.type === 'target_element'){
        if (this.offset) sheet.getRange(...this.cell, ...this.offset).merge();
        range.setValue(this.value);
      }
      if (this.type === 'label_element'){
        if (this.offset) sheet.getRange(...this.cell, ...this.offset).merge();
        range.setValue(this.value);
      }
      if (this.type === 'frame_element'){
        if (this.extent) 
          range = sheet.getRange(...this.cell, ...this.extent);
      }

      for (const prop in this.style)
        Element.setProperty(range, prop, this.style[prop]);

      return this; 
    }

  }
  // assign class static variables
  Element._typesProps = new Map();
  Element._typesProps.set('target_element',
    ['cell', 'offset', 'value', 'style']);
  Element._typesProps.set('label_element',
    ['cell', 'offset', 'value', 'style']);
  Element._typesProps.set('frame_element',
    ['cell', 'extent', 'style']);
  Element._supportedTypes = Array.from(Element._typesProps.keys());


  /**
   * 
   * @property {Array} dayValues - of {Map} record; day records retrieved by date key from records
   */
  class DailyReport {
    /**
     * @param {string} date
     * @param {Map} company
     * @param {Map} dataRecords - reference to all values
     * @param {Function} calculateBalance - takes {string} date as arg and calculates balance till date
     */
    constructor(date, company, dataRecords, calculateBalance){
      this.date = new Date(date);
      this.company = company;
      this.dayValues = dataRecords.get(date);

      this.prevDateStr = ((today=this.date) => {
        const date = new Date(today);
        date.setDate(today.getDate() -1);
        return date.toJSON();
      })();

      this.previous_balance = calculateBalance(this.prevDateStr);
      
      // if balance is negative, that means you spent cash money you didn't collect
      if (this.previous_balance < 0){
        const localPrevDate = new Date(this.prevDateStr).toLocaleDateString('ro-RO');
        throw new Error(
          `Previous day (${localPrevDate}) balance cannot be negative (${this.previous_balance}).`
        ); 
      }

      const [total_input, total_day_output] = this.dayValues.reduce(
        (in_out, record) => {
          if (record.get('I_O_type') === 1){
            in_out[0] += record.get('value');
          } else if (record.get('I_O_type') === 0) {
            in_out[1] += record.get('value');
          }
          return in_out;
        }
      ,[this.previous_balance, 0]);

      this.total_input = total_input; 
      this.total_day_output = total_day_output; 
      this.day_balance = total_input - total_day_output;
    }
    
    setColumnWidths(sheet, widths){
      widths.map((w, i) => sheet.setColumnWidth(i+1, w));
    }
    setRowsHeight(sheet, numRows, height){
      sheet.setRowHeights(1, numRows, height);
    }

    /**
     * Utility function to convert {Object} template to a tree of {Map}.
     * 
     * @param {Object} obj - JSON-like object
     * @param {Array} leafKeys - array of {string}, keys in obj that stores leaves 
     * @returns {Array} [mapTree, leaves] - 
     *   {Map} mapTree - a tree of {Map} instances, having {Element} leaves 
     *   {Map} leaves - reference to every leaf, having keys {Array} [x,y],
     *   that corresponds to 'cell' property of {Element}
     */
    static objToMap(obj, leafKeys, leaves=new Map()){
      const mapTree = new Map();

      for (const key in obj){
        
        if (leafKeys.includes(key)){
          // if is a leaf
          const element = new Element(key, obj[key]);
          mapTree.set(key, element);
          leaves.set(element.keyCell, element); 
        } 
        // does not recurses on arrays; are passed over
        else if (isObject(obj[key])){
          const [subTree, _leaves]  = DailyReport.objToMap(obj[key], leafKeys, leaves);
          mapTree.set(key, subTree);
        }
          
      }
      return [mapTree, leaves]
    }

    render(toSheet, template){

      if (!this.dayValues)
        throw new TypeError(`Cannot render {DailyReport} instance if data values is ${this.dayValues}`);
      
      toSheet.setName(this.date.toLocaleDateString('ro-RO'));
      this.setColumnWidths(toSheet, template._columnWidths);
      const numRows = template._layoutRange[2];
      this.setRowsHeight(toSheet, numRows, template._rowHeight);
      toSheet.getRange(...template._layoutRange).clear();

      const leafKeys = Element.getSupportedTypes();
      // {Map} tree - having {Element} leaves
      // {Map} elements - having key=element.keyCell, and value is {Element} leaf 
      const [tree, elements] = DailyReport.objToMap(template, leafKeys);
      // populate headers (general info displayed on top of report sheet)
      tree.get('companyName').get('target_element').value = company.get('name');
      tree.get('tax_id').get('target_element').value = company.get('tax_id');
      tree.get('reg_num').get('target_element').value = company.get('reg_num');
      

      ((group=tree.get('previous_balance')) => {
        const label = group.get('label_element');
        const target = group.get('target_element');
        // change label according to date (if date is 1st or not)
        if (this.date.getDate() === 1)
          label.value = label.value.replace(/\/ziua/i, '');
        else 
          label.value = label.value.replace(/luna\//i, '');
        target.value = this.previous_balance;
      })();

      const numRecords = this.dayValues.length;

      // {Map} record
      for (const record of this.dayValues){
        
        for (const [parentKey, elementType] of tree.get('record')){
          const defaultElement = elementType.get('target_element');
          const newRecElem = new Element('target_element', defaultElement.templateElement);
          // make a copy (a new record element)
          // writing corresponding values from data record (dayValues)
          if (parentKey === 'date') 
            newRecElem.value = new Date(record.get('date')).toLocaleDateString('ro-RO');
          if (parentKey === 'ref')
            newRecElem.value = record.get('ref');
          if (parentKey === 'doc_type')
            newRecElem.value = record.get('doc_type');
          if (parentKey === 'descr')
            newRecElem.value = record.get('descr'); 
          if (parentKey === 'input')
            newRecElem.value = record.get('I_O_type') === 1 ? record.get('value') : '';
          if (parentKey === 'output')
            newRecElem.value = record.get('I_O_type') === 0 ? record.get('value') : '';
          //updating cell position
          newRecElem.cell[0] = defaultElement.cell[0];
          //updating cell position
          defaultElement.cell[0] += 1;
          // push new element (replacing existing key)
          elements.set(newRecElem.keyCell, newRecElem);
          // push updated key
          elements.set(defaultElement.keyCell, defaultElement);
        }

      }

      // expand body frame accordingly
      ((group=tree.get('body')) => {
        const frame = group.get('frame_element');
        frame.extent[0] += numRecords;
      })();

      ((group=tree.get('day_total_label')) => {
        const label = group.get('label_element');
        // replace '{}' with date in corresponding labels
        label.value = replaceCurly(label.value, this.date.toLocaleDateString('ro-RO')); 
        label.cell[0] += numRecords;
        elements.set(label.keyCell, label);
      })();

      ((group=tree.get('day_total_targets')) => {
        const total_input = group.get('total_input').get('target_element');
        total_input.cell[0] += numRecords;
        total_input.value = this.total_input;
        elements.set(total_input.keyCell, total_input);

        const total_day_output = group.get('total_day_output').get('target_element');
        total_day_output.cell[0] += numRecords;
        total_day_output.value = this.total_day_output;
        elements.set(total_day_output.keyCell, total_day_output);
      })();

      ((group=tree.get('day_balance')) => {
        const label = group.get('label_element');
        label.value = replaceCurly(label.value, this.date.toLocaleDateString('ro-RO')); 
        label.cell[0] += numRecords;
        elements.set(label.keyCell, label);
        const target = group.get('target_element');
        target.cell[0] += numRecords;
        target.value = this.day_balance;
        elements.set(target.keyCell, target);
      })();


      // render all elements that has a value 
      const renderedElements = new Map();
      for (const [key, element] of elements){
        const rendered = element.render(toSheet);
        renderedElements.set(key, rendered);
        }

      return renderedElements;
    }
  }


  class Report{
    constructor(fromDate, toDate, company, dataRecords, template){
      this.fromDate = fromDate;
      this.toDate = toDate;
      this.company = company;
      this.dataRecords = dataRecords;
      const recordDates = Array.from(dataRecords.keys());
      recordDates.sort();
      this.recordDates = recordDates;
      this.template = template;
    }

    /**
     * returns a closure to calculate balance till date
     */
    balanceCalculator(){
      // closure variables
      const sortedDates = this.recordDates;
      const dataRecords = this.dataRecords;

      return(
        (currentDateStr) => {
          if (! typeof currnetDateStr === 'string')
            throw new TypeError(
              `Typeof currentDateStr is ${typeof currentDateStr}. Expected string.`);
          if (isNaN(new Date(currentDateStr)))
            throw new TypeError(
              `${typeof currentDateStr} ${currentDateStr} is not a valid date JSON string.`);

          let total = 0;
          for (const dateStr of sortedDates){
            if (currentDateStr < dateStr)
              return total;
            const dayRecords = dataRecords.get(dateStr);
            total += dayRecords.reduce((dayTotal, record) => {

              const recordValue = record.get('value');
              if (typeof recordValue !== 'number')
                throw new TypeError(
                  `Expected Number! Typeof record value ${recordValue}: ${typeof recordValue}. `+
                  `Date key: ${dateStr}. Local Date: ${new Date(dateStr).toLocaleDateString('ro-RO')}`);

              if (record.get('I_O_type') === 1)
                return dayTotal + recordValue;
              if (record.get('I_O_type') === 0)
                return dayTotal - recordValue;
          }, 0);
        }
        return total;
        }
      );
    }
    
    render(targetSpreadsheet, template){
      
      /* for every date between fromDate and toDate:
       *   collect dataRecords and group by date in a {Map},
       *   generate an instance of {DailyReport},
       *   create a new {Sheet} instance in {Spreadsheet} and name it with date,
       *   render every dayReport to sheet according with date,
       *   and DONE
       */
          
      // delete existing sheets except first
      v>1 && addMessage(`Deleting existing report sheets except first`);
      targetSpreadsheet.getSheets().forEach(sheet =>{
        if (sheet.getIndex() === 1) 
          // cover sheet 
          sheet.setName('Cover');
        else
          targetSpreadsheet.deleteSheet(sheet);
      }
      );
      const dates = datesBetween(fromDate, toDate);

      v>1 && addMessage(`Rendering reports between ${fromDate.toLocaleDateString('ro-RO')} and ${toDate.toLocaleDateString('ro-RO')}`);

      let sheetIndex = 1
      for (const date of dates){
        const dayTrades = dataRecords.get(date);
        if (!dayTrades) continue;
        const sheet = targetSpreadsheet.insertSheet(sheetIndex++);
        const dayReport = new DailyReport(date, company, dataRecords, this.balanceCalculator());
        dayReport.render(sheet, this.template);
        v>1 && addMessage(`Day report ${new Date(date).toLocaleDateString('ro-RO')} rendered!`);
      }

      return;
    }
  }

  //----------------------------------------------------------------
  //-------------- render all reports ------------------------------
  const report = new Report(fromDate, toDate, company, dataRecords, template)
  report.render(targetSpreadsheet);
  //================================================================

  v > 0 && addMessage('Procedure renderReport END');
} // renderReport END

/**
 * Searches for company data records in provided links to standalone spreadsheets,
 * and populate corresponding rawDataSheet.
 *
 * @param {Map} company - dict with company info keys like 'name', 'alias', etc
 * @param {Iterable} dataLinks - records with links (urls) of google sheets
 *      - {string} like 'https://docs.google.com/spreadsheets/d/<< sheetId >>/edit#gid=xxxxxxxxxx';
 */
function importData(
  fromDate, toDate, company, dataLinks, identifierPattern, sheetToImportTo, SpreadsheetApp){

  const getRecords = libraryGet('getRecords');
  const addMessages = libraryGet('addMessages');

  // initialize debug messaging
  const thisProcedure = importData;
  const addMessage = addMessages(thisProcedure); // adds prop {Map} messages
  // @prop {number} verbosity could be added at procedure call
  const v = thisProcedure.verbosity ? thisProcedure.verbosity : 0;

  v>0 && addMessage('Procedure importData begin');
  v>1 && addMessage(`Company alias: ${company.get('alias')}`);

  // tableName the prefix before '.' in field name, like tableName.fieldName
  const linkTableName = 'link';

  // list of google sheets ids 
  const sheetIds = (dataLinks => {
    const ids = [];
    for (const link of dataLinks){
      if (link)
        try {
          const sheetId = extractId(link);
          ids.push(sheetId);
          numOfEmpty = 0;
        } catch(e){
          addMessage(`${e}\nSeems that link:\n${link}\ndoes not match pattern.`);
        }
      else throw new ReferenceError('No spreadsheet link.');
    }
    return ids;
  })(dataLinks)
  v>0 && addMessage(`Found ${sheetIds.length} ids in links`)
  v>1 && addMessage(`sheeIds: ${sheetIds}`);

  // list of source Spreadsheets opened by ids;
  const srcSpreadsheets = sheetIds.reduce(
    (spreadsheets, sheetId)  => {
      try{
        const spreadsheet = SpreadsheetApp.openById(sheetId);
        spreadsheets.push(spreadsheet);
        return spreadsheets;
      } catch(e){
        addMessage(`When opening sheet with id ${sheetId}\ngot ${e}`);
        return spreadsheets;
      }
    }, []
  );

  if ( ! srcSpreadsheets.length){
    addMessage('No source spreadsheets opened!'); 
    return 2;
  } else {
    v>0 && addMessage(`Spreadsheets opened ${srcSpreadsheets.length}, [${srcSpreadsheets.map(ss => ss.getName())}]`);
  }

  const foundRecords = new Map(); 
  for (const sheet of srcSpreadsheets){
    for (const [dateStr, record] of searchRecords(sheet, identifierPattern)){
      foundRecords.set(dateStr, record);
    }
  }

  if (!foundRecords.size)
    v>0 && addMessage(
      `No records found in spreadsheet ${srcSpreadsheets[0].getName()}`);
  else
    v>1 && addMessage(
      `Found ${foundRecords.size} day-records in spreadsheets: [${srcSpreadsheets.map(ss => ss.getName())}]`);

  
  // retrieve existing records in raw data sheet
  const existingRecords = getRecords(sheetToImportTo);
  v>1 && addMessage(`${existingRecords.size} day-records exists in ${sheetToImportTo.getName()}`);

  const dates = datesBetween(fromDate, toDate);
  v>1 && addMessage(`Searching found-records between ${new Date(fromDate).toLocaleDateString('ro-RO')} and ${new Date(toDate).toLocaleDateString('ro-RO')}...`);

  for (const dateStr of dates){
    const foundDateRecords = foundRecords.get(dateStr);
    const existingDateRecords = existingRecords.get(dateStr);
    if (foundDateRecords && existingDateRecords){
      v>0 && addMessage(`Duplicates found on date ${new Date(dateStr).toLocaleDateString('ro-RO')}.`);
      v>1 && addMessage('resolving same-date-key conflicts...');
      const mergedDateRecords = mergeDateRecords(foundDateRecords, existingDateRecords);
      existingRecords.set(dateStr, mergedDateRecords);
    } else if (foundDateRecords){
      existingRecords.set(dateStr, foundDateRecords);
    }
  }

  v>1 && addMessage('updating raw data sheet...')
  const rawValues = [];
  const keyDates = Array.from(existingRecords.keys()).sort();
  keyDates.forEach(
    dateStr => {
      for (const record of existingRecords.get(dateStr)){ 
        rawValues.push(
          [new Date(dateStr),
          record.get('ref'),
          record.get('doc_type'),
          record.get('descr'),
          record.get('I_O_type'),
          record.get('value')]
        );
      }
    }
  );

  const rawDataRange = sheetToImportTo.getRange(2, 1, rawValues.length, rawValues[0].length);
  // delete all existing records
  sheetToImportTo.getRange('A2:F').clear();
  v>1 && addMessage(`Deleted all 'A2:F' values from sheet ${sheetToImportTo.getName()}!`); 
  // writing new values
  
  v>1 && addMessage(`Writing new values...`);
  rawDataRange.setValues(rawValues);

  v>0 && addMessage('Procedure importData END');
} // importData END

/**
 * @param {Date} fromDate
 * @param {Date} toDate
 * @param {Map} company - dict with company info keys like 'name', 'alias', etc
 * @param {Sheet} rawDataSheet
 */
function cleanRawData(fromDate, toDate, company, rawDataSheet){

  const getFieldNames = libraryGet('getFieldNames');
  const FieldValidator = libraryGet('FieldValidator');
  const getType = libraryGet('getType');
  const addMessages = libraryGet('addMessages');

  // initialize debug messaging
  const thisProcedure = cleanRawData;
  const addMessage = addMessages(thisProcedure); // adds prop {Map} messages
  // @prop {number} verbosity could be added at procedure call
  const v = thisProcedure.verbosity ? thisProcedure.verbosity : 0;

  v>0 && addMessage('Procedure cleanRawData begin');
  
  // retrieve from spreadsheet
  const dataRange = rawDataSheet.getRange('A1:Z');
  const values = dataRange.getValues();
  
  v>1 && addMessage(`values retrieved from sheet ${rawDataSheet.getName()}`);

  const startTimer = Date.now();

  const fieldNames = getFieldNames(values[0]);
  v>1 && addMessage(`fieldNames are ${JSON.stringify(fieldNames)}`);
  // list of all indexes that have an associated field
  const fieldIndexes = [];
  for (const fieldName in fieldNames)
    fieldIndexes.push(fieldNames[fieldName]);
  
  const allIndexes = Array(values[0].length).fill(0).map((e,i)=>e+i);
  const nonFieldIndexes = allIndexes.filter(i => ! fieldIndexes.includes(i));

  // field descriptions
  const validator = new FieldValidator();
  for (const fieldName in fieldNames){
    const fieldIndex = fieldNames[fieldName];
    if (fieldName === 'date')
      validator.setField(fieldName,fieldIndex,'Date');
    else if (fieldName === 'ref')
      validator.setField(fieldName,fieldIndex,'string');
    else if (fieldName === 'doc_type')
      validator.setField(fieldName,fieldIndex,'string');
    else if (fieldName === 'descr')
      validator.setField(fieldName,fieldIndex,'string');
    else if (fieldName === 'I_O_type')
      validator.setField(fieldName,fieldIndex,'number',0,1);
    else if (fieldName === 'value')
      validator.setField(fieldName,fieldIndex,'number');
    else
      throw new Error(`Unknown fieldName ${fieldName}`);
  }

  const convertType = (recordVal, newType) => {
    const currentType = getType(recordVal);
    if (currentType === newType)
      return recordVal;
    
    const typeConstructors = new Map();
    typeConstructors.set('number', Number);
    typeConstructors.set('string', JSON.stringify);
    typeConstructors.set('Date', dateLike => new Date(dateLike));

    try {
      const converted = typeConstructors.get(newType)(recordVal);
      return converted;
    } catch(e) {
      // if we got here, it means that type of record value was not converted
      const err = new Error('Type not converted');
      err.method = 'convertType';
      err.recordVal = recordVal;
      err.currentType = currentType;
      err.expectedType = newType;
      err.orginalError = e;
      throw err;
    }
  };
  
  // this is a set of unique strings (hash of record)
  // I used array for performance of "includes" and "indexOf" (node.js v14.8.0)
  const uniques = []; 
  // store index of unique record
  const indexesOfUnique = [0]; // index 0 is for fieldNames
  
  let emptyRowCount = 0;
  let row_i = 0;
  while(++row_i < values.length){
    const record = values[row_i];

    // if 10 empty records are encountered then is end of data set
    if (emptyRowCount > 9){
      v>1 && addMessage(`found ${emptyRowCount} empty rows, so break`); 
      break;
    }
    const rowIsEmpty = record.reduce((isEmpty,val)=>{
      return [NaN,'',null,undefined].includes(val) ? isEmpty : false;
    }, true);
    if (rowIsEmpty){ ++emptyRowCount;
      continue;
    }
    // if got here, then emptyRowCount < 10, so reset
    emptyRowCount = 0;

    // delete all values from fields with indexes that are not in fieldIndexes
    record.forEach((v, i) => {
      if (nonFieldIndexes.includes(i))
        delete record[i];
    });

    // validation begin
    for (const fieldName in fieldNames){
      // fieldIndex should correspond with index of value from record
      const fieldIndex = fieldNames[fieldName]; 
      const testValue = record[fieldIndex];

      try { validator.validate(fieldName,testValue);}
      catch(e) {
        if (getType(e) === 'TypeError'){
          if ( ! e.expectedType) throw e;
          // try to convert according with field type (e.expectedType)
          const converted = convertType(testValue, e.expectedType);

          // validate again, if not thows then is a correct value/type
          try { validator.validate(fieldName, converted);}
          catch(e){ e.rowIndex = row_i; throw e}

          v>1 && addMessage(`converted {${getType(testValue)}} ${testValue} `+
          `to {${getType(converted)}} ${converted} in row_i=${row_i}`);

          // write value again in record array
          record[fieldIndex] = converted;
        } else {
          e.rowIndex = row_i;
          throw e;
        }
      }
    }
    // now the record should be of correct type for every value
    // in order to check if record is duplicate, hash it
    // and add to a set of unique values
    const recordHash = JSON.stringify(record); 
    // if is already in uniques, then is duplicate
    if ( ! uniques.includes(recordHash)){
      uniques.push(recordHash);
      indexesOfUnique.push(row_i);
    }
  }
  
  v>1 && addMessage(`values.length ${values.length}`)
  v>1 && addMessage(`uniques.length ${uniques.length}`)

  // now we got all indexes of unique values 
  // let't remove them by constructing a new set of values
  const newValues = indexesOfUnique.map(i => values[i]);
  v>1 && addMessage(`newValues.length ${newValues.length} <- this includes first row - fields names`);

  // sort records by date except first row (index 0) - field names
  // in order to do that, temporary change first record value (first field name)
  // such that this remains the first row after sort
  const firstFieldName = newValues[0][0];
  newValues[0][0] = new Date(0);
  newValues.sort((rec1, rec2) => rec1[0] - rec2[0]); 
  newValues[0][0] = firstFieldName;
  v>1 && addMessage(`newValues sorted`);

  // enlarge newValues with empty values to match range
  const limit = values.length - newValues.length; 
  let start_i = newValues.length;
  while (start_i < values.length){
    newValues.push(Array(values[0].length));
    ++start_i;
  }
  
  v>1 && addMessage(`procedure done in ${(Date.now() - startTimer)/1000} sec`);
  // reset values on range
  dataRange.setValues(newValues); 
  v>1 && addMessage('all new values written');
  v>0 && addMessage('procedure cleanRawData END');
  return 'done';
} // procedure cleanRawData END


/**
 * Arguments validator
 */
function argumentsValidator(){
  // variable enclosed by setArgTypes function
  const _argTypes = [];

  const setArgTypes = (...argTypes)=>argTypes.forEach(
    type =>{
      if (typeof type !== 'string')
        throw new TypeError(`Type descriptor ${typeof type} is not valid`);
      _argTypes.push(type);
      }
  );
  
  const validateArgs = (...currArgs) => {
    if (currArgs.length !== _argTypes.length){
      const err = new TypeError('Wrong number of arguments');
      err.currArgsLength = currArgs.length;
      err.expectedArgsLength = _argTypes.length;
      err.currArgsTypes = currArgs;
      err.expectedArgsTypes = _argTypes; 
      throw err;
    }
    _argTypes.forEach((type, i) =>{
      const currArg = currArgs[i];
      if (getType(currArg) !== type){
        const err = new TypeError('Invalid Signature');
        err.expectedType = _argTypes[i];
        throw err;
      }
    });
  }; 
  // returns setter and validator closure functions
  return [setArgTypes, validateArgs];
} // argumentsValidator END

/**
 * @param {*} obj
 * @returns {string}
 */
function getType(obj){
  if (obj === null)
    return 'null';
  if (obj === undefined)
    return 'undefined';
  if (Object.is(obj, NaN))
    return 'nan';
  if (Object.is(obj, Boolean(obj)))
    return 'boolean';
  if (typeof obj === 'object')
    // returns 'Object', 'Array', 'Map', 'Set', etc
    return obj.constructor.name;
  if (typeof obj === 'function')
    // returns 'Function'
    return obj.constructor.name;
  // returns 'number', 'string'
  return typeof obj;
}


class FieldValidator{
  constructor(){
    this._fieldNames = new Map();
    this._fieldIndexes = new Map();
    this._validTypes = ['string', 'number', 'boolean', 'Date'];
  }
  /**
   * @param {string} fieldName
   * @param {number} fieldIndex
   * @param {string} fieldType
   * @param {*} [minValue]
   * @param {*} [maxValue]
   * @param {Set} [exactValues]
   */
  setField(fieldName, fieldIndex, fieldType, minValue, maxValue, exactValues,sentinel){
    
    const [setArgTypes, validateArgs] = argumentsValidator();
    
    // setField is called with these arguments
    const currArgs = [
      fieldName, fieldIndex, fieldType, minValue, maxValue, exactValues,];
    // sentinel should be undefined to limit the number of arguments
    if (sentinel !== undefined) currArgs.push('wrong argument');

    // dynamic arguments types
    
    const null_undefined = ['null','undefined'];
    const isNone = typeStr => null_undefined.includes(typeStr);

    // required minValue type
    const minValueArgType = getType(minValue);
    const minValueType = isNone(minValueArgType) ?
      minValueArgType : fieldType;

    // required maxValue type
    const maxValueArgType = getType(maxValue);
    const maxValueType = isNone(maxValueArgType) ?
      maxValueArgType : fieldType;

    // required exactValues type
    const exactValuesArgType = getType(exactValues);
    // if both minValue and maxValue are null/undefined
    // then exactValues can be {Set} or null/undefined
    // else if either minValue or maxValue are not null/undefined
    // thin exactValues can be only null/undefined
    if (isNone(minValueType) && isNone(maxValueType)){
      if (isNone(exactValuesArgType))
        setArgTypes(
          'string','number','string',minValueType,maxValueType,exactValuesArgType,);
      else
        setArgTypes(
          'string','number','string',minValueType,maxValueType,'Set');
    } else {
      // here either minValueType or maxValueType are not none
      if (isNone(exactValuesArgType))
        setArgTypes(
          'string','number','string',minValueType,maxValueType,exactValuesArgType);
      else // here we got not none type for last arg (exactValues)
        // chosen default type for exactValues is 'undefined'
        setArgTypes(
          'string','number','string',minValueType,maxValueType,'undefined');
    }
    // validate all arguments ==================
    validateArgs(...currArgs);

    const field = new Map();
    if (this._fieldNames.has(fieldName))
      throw new Error('fieldName already exists');
    if (this._fieldIndexes.has(fieldIndex))
      throw new Error('fieldIndex already exists'); 
    field.set('name', fieldName);
    field.set('index', fieldIndex);
    field.set('type', fieldType);

    if (minValue !== null && minValue !== undefined && exactValues === undefined)
      field.set('minValue', minValue);
    // set max value only if exactValues are not proviced
    if (maxValue !== null && maxValue !== undefined && exactValues === undefined)
      field.set('maxValue', maxValue);
    // throw if type of exactValues is not the same as fieldType
    if (exactValues !== null && exactValues !== undefined){
      for (const exactVal of exactValues.keys()){
        if (getType(exactVal) !== fieldType){
          const err = new TypeError('Invalid field exactType');
          err.fieldName = fieldName;
          err.expectedType = fieldType;
          err.currentType = getType(exactVal);
          throw err;
        }
      }
      field.set('exactValues', exactValues);
    }

    this._fieldNames.set(fieldName, field); 
    // cannot be tow fields with same index;
    this._fieldIndexes.set(fieldIndex, field);
    return this;
  }
  
  /**
   * @param {string} fieldName
   * @param {*} testValue
   */
  validate(fieldName, testValue){
    const [setArgTypes, validateArgs] = argumentsValidator();
    // validate arguments
    setArgTypes('string', getType(testValue));
    validateArgs(...arguments);

    class ValueError extends Error{
      constructor(msg){
        super(msg);
        this.name = 'ValueError';}}
    class NotFoundError extends Error{
      constructor(msg){
        super(msg);
        this.name = 'NotFoundError';}}

    if ( ! this._fieldNames.has(fieldName)){
      const err = new NotFoundError('Field not found');
      err.fieldName = fieldName;
      throw err;
    }

    const field = this._fieldNames.get(fieldName);

    const valType = getType(testValue); 
    // test for type
    if (field.get('type') !== valType){
      const err = new TypeError('Invalid Value Type');
      err.fieldName = fieldName;
      err.expectedType = field.get('type');
      err.currentType = valType;
      throw err;
    }

    // test for value
    if(field.get('type') === 'Date'){
      // explicit test for invalid date
      if (Object.is(testValue.getTime(), NaN)){
        const err = new ValueError('Invalid Date Value');
        err.fieldName = fieldName;
        err.expectedValue = `something like ${new Date()}`;
        err.currentValue = testValue;
        throw err;
      }
    }
    if (field.has('minValue')){
      if (testValue < field.get('minValue')){
        const err = new ValueError('Value less than minValue');
        err.fieldName = fieldName;
        err.expectedValue = `>= ${field.get('minValue')}`;
        err.currentValue = `${testValue}`;
        throw err;
      }
    }
    if (field.has('maxValue')){
      if (testValue > field.get('maxValue')){
        const err = new ValueError('Value greater than maxValue');
        err.fieldName = fieldName;
        err.expectedValue = `<= ${field.get('maxValue')}`;
        err.currentValue = testValue;
        throw err;
      }
    }
    if (field.has('exactValues')){
      if ( ! field.get('exactValues').has(testValue)){
        const err = new ValueError('Value not found in exactValues Set');
        err.fieldName = fieldName;
        err.currentValue = testValue;
        throw err;
      }
    }
    
    return true;
  }
  getFieldByName(fieldName){
    return this._fieldNames.get(fieldName);
  }
  getFieldByIndex(fieldIndex){
    return this._fieldIndexes.get(fieldIndex);
  }
}


/**
 * @param {Array} record
 * @param {Object} fieldNames - key: {string} fieldName, val: {Number} index
 */
function validateRecord(record, fieldNames){

  const validators = new Map();

  validators.set('date',
    date => {
      if (isNaN(date.getTime()))
        throw new TypeError(
          `In field ${fieldNames['date']} ${typeof date} ${date} is not valid Date.`);
    });

  validators.set('ref',
    (ref, fieldName) => {
      const type = typeof ref;
      const val = ref;
      const expectedType = 'string';
      const errorInfo = {
          field_name: fieldName,
          field_index: fieldNames[fieldName],
          value_type: type,
          value: val,
          expected_type: expectedType,
          errors: [],
          fromRef: "fromRef"
      };

      if (type !== expectedType)
        errorInfo.errors.push('type_error');
        //throw new TypeError(JSON.stringify({...errorInfo, )); 

      if (errorInfo.errors.length)
        throw new Error(JSON.stringify(errorInfo));
    });

  validators.set('doc_type',
    doc_type => {
      const type = typeof doc_type;
      const val = doc_type;
      const expectedType = 'string';

      if (type !== expectedType)
        throw new TypeError(JSON.stringify({
          field_name: fieldName,
          field_index: fieldNames[fieldName],
          value_type: type,
          value: val,
          expected_type: expectedType
        })); 
    });
  validators.set('descr',
    descr => {
      const type = typeof descr;
      const val = descr;
      const expectedType= 'string';

      if (type !== expectedType)
        throw new TypeError(JSON.stringify({
          field_name: fieldName,
          field_index: fieldNames[fieldName],
          value_type: type,
          value: val,
          expected_type: expectedType
        })); 
    });
  validators.set('I_O_type',
    I_O_type => {
      const type = typeof I_O_type;
      const val = I_O_type;
      const expectedType = 'number';
      if (type !== expectedType)
        throw new TypeError(JSON.stringify({
          field_name: fieldName,
          field_index: fieldNames[fieldName],
          value_type: type,
          value: val,
          expected_type: expectedType
        })); 
      if (val < 0 || val > 1)
        throw new TypeError(
          `In field ${fieldNames['I_O_type']} ${type}, val:${val} is not 1 or 0`); 
    });
  validators.set('value',
    value => {
      const type = typeof value;
      const val = value;
      const expectedType = 'number';
      if (type !== expectedType)
        throw new TypeError(JSON.stringify({
          field_name: fieldName,
          field_index: fieldNames[fieldName],
          value_type: type,
          value: val,
          expected_type: expectedType
        })); 
    });

  // validate record (row)
  for (const fieldName in fieldNames){
    const i = fieldNames[fieldName];
    const val = record[i];
    if ( ! validators.has(fieldName))
      throw new ReferenceError(`Validator function not set for field ${fieldName}`);
    validators.get(fieldName)(val, fieldName);
  }
  
} // function validateRecord END

function getFieldNames(firstRow){
  // object to store fieldName as key and associated index as value
  // this object will be queried for every value and it seems JS is very fast at
  // reading plain objects (key is string and value does not change)
  // it seems it is better over a {Map} instance;
  const fieldNames = firstRow.reduce(
    (indexes, fieldName, i) => {
      // only fields that has a value are recorded
      if (fieldName) indexes[fieldName] = i;
      return indexes;
    },
    Object.create(null));

  return fieldNames;
}

/**
 * @param {Sheet} rawDataSheet
 * @param {Object[]} fieldDescriptors
 * 
 * @return {Map} records - {string} key date.toJSON(), {Array} values
 * 
 * returns something like ....
 */
function getRecords(rawDataSheet, fieldDescriptors){
  //const validate = libraryGet('validateRecord');
  const getFieldNames = libraryGet('getFieldNames');
  const FieldValidator = libraryGet('FieldValidator');
  const addMessages = libraryGet('addMessages');

  // add default value to fieldDescriptors
  if (fieldDescriptors===undefined)
    fieldDescriptors = [
      {fieldName:"date",fieldType:"Date"},
      {fieldName:"ref",fieldType:"string"},
      {fieldName:"doc_type",fieldType:"string"},
      {fieldName:"descr",fieldType:"string"},
      {fieldName:"I_O_type",fieldType:"number",minValue:0,maxValue:1},
      {fieldName:"value",fieldType:"number"}
    ];

  // initialize debug messaging
  const thisProcedure = getRecords;
  const addMessage = addMessages(thisProcedure); // adds prop {Map} messages
  // @prop {number} verbosity could be added at procedure call
  const v = thisProcedure.verbosity ? thisProcedure.verbosity : 0;

  const range = rawDataSheet.getRange('A1:Z');
  const values = range.getValues();
  v>0 && addMessage(`Records in ${rawDataSheet.getName()}: ${values.length}`);
  
  // these are the names found on first row of sheet;
  const fieldNames = getFieldNames(values[0]);
  const expectedFieldNames = [];

  const validator = new FieldValidator();
  for (const fieldDescr of fieldDescriptors){
    const fieldName = fieldDescr.fieldName;
    expectedFieldNames.push(fieldName);
    if (fieldName in fieldNames){
      const fieldIndex = fieldNames[fieldName];
      validator.setField(
        fieldName, fieldIndex, fieldDescr.fieldType,
        fieldDescr.minValue,
        fieldDescr.maxValue,
        fieldDescr.exactValues);
    } else {
      const err = new TypeError('Not found');
      err.fieldName = fieldName;
      throw err;
    }
  }

  const records = new Map();
  let emptyRowCount = 0;

  for (let row_i=1; row_i < values.length; row_i++){
    const row = values[row_i];

    // if 10 empty records are encountered then is end of data set
    if (emptyRowCount > 9){
      v>1 && addMessage(`found ${emptyRowCount} empty rows, so break`); 
      break;
    }
    const rowIsEmpty = row.reduce((isEmpty,val)=>{
      return [NaN,'',null,undefined].includes(val) ? isEmpty : false;
    }, true);
    if (rowIsEmpty){ ++emptyRowCount;
      continue;
    }
    // if got here, then emptyRowCount < 10, so reset
    emptyRowCount = 0;

    const record = new Map();

    let dateKey = '';

    for (const fieldName of expectedFieldNames){
      const field_i = fieldNames[fieldName];
      const thisValue = row[field_i];

      try { validator.validate(fieldName, thisValue);}
      catch(err){
        err.sheetName = rawDataSheet.getName();
        err.row = row_i + 1;
        err.thisProcedureName = thisProcedure.name;
        throw err;
      }
      if (fieldName==='date'){
        dateKey = thisValue.toJSON();
        record.set(fieldName, dateKey); 
      } else {
        record.set(fieldName, thisValue); 
      }
    }
    // a record will be retrieved by date
    records.has(dateKey) && records.get(dateKey).push(record)
      || records.set(dateKey, [record]);
  }

  return records;
} // function getRecords END

const libScopeVar = '_libscopevar';
function dummy_func(ceva){
  return ceva + libScopeVar + fromMain;
}

library.set('settings', Settings);
library.set('renderReport', renderReport);
library.set('importData', importData);
library.set('cleanRawData', cleanRawData);
library.set('validateRecord', validateRecord);
library.set('getFieldNames', getFieldNames);
library.set('getRecords', getRecords);
library.set('FieldValidator', FieldValidator);
library.set('addMessages', addMessages);
library.set('Log', Log);
library.set('argumentsValidator', argumentsValidator);
library.set('getType', getType);
library.set('dummy_func', dummy_func);

// returns required function or class
return library.get(required);
} // function libraryGet END


