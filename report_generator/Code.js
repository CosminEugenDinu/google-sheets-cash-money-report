
/**
 * Main function, call from sheet button, where it is assigned
 * This function is considered global
 */
function makeReport() {

// global settings variables
const REPORT_GENERATOR_SPREADSHEET_ID = "1nDNPcpgP9TlAcTSKASwFxuQQswdAI7Wm3I8Z-7rSGnE";
const REPORT_SPREADSHEET_ID = "1e0nIxg2pNLnnSPmKdmkRHS7cl7kVEj2G9CKBksXflEk";
const INTERFACE_SHEET_NAME = "Interface";
const SETTINGS_SHEET_NAME = "settings";
const RAWDATA_SHEET_SUFFIX = "_rawdata";

const TEMPLATE = defaultTemplate();

// instantiate log function
const log = Log(REPORT_GENERATOR_SPREADSHEET_ID, 0, [10,5]);
  
const repGenSprSheet = SpreadsheetApp.openById(REPORT_GENERATOR_SPREADSHEET_ID);
const repSprSheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);

// interface is a {Sheet} with selects and button that triggers 'makeReport'
const interface = repGenSprSheet.getSheetByName(INTERFACE_SHEET_NAME);
const settings = repGenSprSheet.getSheetByName(SETTINGS_SHEET_NAME);
 
const companies = getCompanies(settings);
const companyAliases = Array.from(companies.keys());
const computedRawDataSheetNames = companyAliases.map(
  alias => alias + RAWDATA_SHEET_SUFFIX); 
const rawDataSheets = repGenSprSheet.getSheets().filter(
  sheet => {
  const sheetName = sheet.getSheetName();
  // get all sheets except Interface and settings
  return (sheetName !== INTERFACE_SHEET_NAME) && (sheetName !== SETTINGS_SHEET_NAME);
});
  
// if company alias name was changed in settings
updateRawDataSheetNames(rawDataSheets, computedRawDataSheetNames);

// user selects in interface
const procedure = interface.getSheetValues(8,1,1,1)[0][0];
const companyAlias = interface.getSheetValues(8,2,1,1)[0][0];
const [fromDate, toDate] = interface.getSheetValues(8,3,1,2)[0];

const verbosity = interface.getSheetValues(8,5,1,1)[0][0];
// v=0 - critical, v=1 - informal, v=2 - too verbose
const v = +verbosity;

// data source sheet corresponds with chosen company alias from drop-down
const srcRawDataSheet = repGenSprSheet.getSheetByName(companyAlias+RAWDATA_SHEET_SUFFIX);
const dataRange = srcRawDataSheet.getRange('A2:F');


const company = companies.get(companyAlias);
const dataRecords = getRecords(dataRange);
const template = TEMPLATE;
const targetSpreadsheet = repSprSheet;

//-----------------------------------------------------------------
procedure === 'renderReport' && renderReport(
  fromDate,
  toDate,
  company,
  dataRecords,
  template,
  targetSpreadsheet);

//-----------------------------------------------------------------

const dataLinks = settings.getSheetValues(1,16,1000,3);
const [sheetToImportTo] = rawDataSheets.filter(
  sheet => sheet.getName() === company.get('alias')+RAWDATA_SHEET_SUFFIX
  )
procedure === 'importData' && importData(
  fromDate,
  toDate,
  company,
  dataLinks,
  sheetToImportTo
);
//-----------------------------------------------------------------



// -------------------------- library --------------------------------

function renderReport(fromDate, toDate, company, dataRecords){
v>0&& log('Procedure renderReport begin');

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
    const jsonDateRe = /^(\d{4}-\d{2}-)(\d{2})(T[\d:.]*Z)$/;
    if ( ! date.match(jsonDateRe))
      throw new TypeError(`JSON date str ${jd} does not match`);
    this.prevDateStr = date.replace(jsonDateRe, (m,g1,day,g3) => g1+(+day-1)+g3);
    this.previous_balance = calculateBalance(this.prevDateStr);
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

    // number of sheet row - retrieved from an record element ('date');
    //let rowNum = tree.get('record').get('date').get('target_element').cell[0] - 1;
    const numRecords = this.dayValues.length;
    const newElements = new Map();
    //let i = 0;
    // {Map} record
    for (const record of this.dayValues){
      //++rowNum
      
      for (const [parentKey, elementType] of tree.get('record')){
        //i++
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

    ((group=tree.get('total')) => {
      const label = group.get('label_element');
      // replace '{}' with date in corresponding labels
      label.value = replaceCurly(label.value, this.date.toLocaleDateString('ro-RO')); 
      label.cell[0] += numRecords;
      elements.set(label.keyCell, label);
      const target = group.get('target_element');
      target.cell[0] += numRecords;
      elements.set(target.keyCell, target);
    })();

    ((group=tree.get('day_balance')) => {
      const label = group.get('label_element');
      label.value = replaceCurly(label.value, this.date.toLocaleDateString('ro-RO')); 
      label.cell[0] += numRecords;
      elements.set(label.keyCell, label);
      const target = group.get('target_element');
      target.cell[0] += numRecords;
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
      let total = 0;
      for (const dateStr of sortedDates){
        if (dateStr > currentDateStr)
          return total;
        const dayRecords = dataRecords.get(dateStr);
        total += dayRecords.reduce((dayTotal, record) => {
          if (record.get('I_O_type') === 1)
            return dayTotal + record.get('value');
          if (record.get('I_O_type') === 0)
            return dayTotal - record.get('value');
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
    v>2&& log(`Deleting existing report sheets except first`);
    targetSpreadsheet.getSheets().forEach(sheet =>{
      if (sheet.getIndex() === 1) 
        // cover sheet 
        sheet.setName('Cover');
      else
        targetSpreadsheet.deleteSheet(sheet);
    }
    );
    const dates = datesBetween(fromDate, toDate);
    v>2&& log(`Rendering reports between ${fromDate.toLocaleDateString('ro-RO')} and ${toDate.toLocaleDateString('ro-RO')}`);

    let sheetIndex = 1
    for (const date of dates){
      const dayTrades = dataRecords.get(date);
      if (!dayTrades) continue;
      const sheet = targetSpreadsheet.insertSheet(sheetIndex++);
      const dayReport = new DailyReport(date, company, dataRecords, this.balanceCalculator());
      dayReport.render(sheet, this.template);
      v>2&& log(`Day report ${new Date(date).toLocaleDateString('ro-RO')} rendered!`);
    }

    return;
  }


}

//----------------------------------------------------------------
//-------------- render all reports ------------------------------
const report = new Report(fromDate, toDate, company, dataRecords, template)
report.render(targetSpreadsheet);
//================================================================

v>0&& log('Procedure renderReport END');
} // renderReport END


/**
 * Searches for company data records in provided links to standalone spreadsheets,
 * and populate corresponding rawDataSheet.
 *
 * @param {Map} company - dict with company info keys like 'name', 'alias', etc
 * @param {Array[][]} dataLinks - records with links (urls) of google sheets
 *      - dataLinks[0]: list of fields names, like [link.company1, link.company2, ...]
 *      - dataLinks[1...n]: records of links, like ['link1', 'link2', ...] 
 *      - dataLinks[row][column]: {string} like 'https://docs.google.com/spreadsheets/d/<< sheetId >>/edit#gid=xxxxxxxxxx';
 */
function importData(fromDate, toDate, company, dataLinks, sheetToImportTo){
  v>0&& log('Procedure importData begin');
  v>2&& log(`Company alias: ${company.get('alias')}.`);

  // tableName the prefix before '.' in field name, like tableName.fieldName
  const linkTableName = 'link';
  const linkFieldNames = dataLinks[0].map(
    fullName => fullName.replace(new RegExp(`^${linkTableName}\.`), '')
    );
  // index of field with name = company alias
  const companyIndex = linkFieldNames.indexOf(company.get('alias'));

  // list of google sheets ids 
  const sheetIds = (dataLinks => {
    const links = [];
    // count the number successive empty rows  
    let numOfEmpty = 0; 
    let i = 1;
    while (numOfEmpty < 10){
      const link = dataLinks[i++][companyIndex];
      if (link){
        try {
          const sheetId = extractId(link);
          links.push(sheetId);
          numOfEmpty = 0;
        } catch(e){
          v>0&& log(`${e}\nSeems that link:\n${link}\ndoes not match pattern.`);
        }
      } else {
        ++ numOfEmpty;
      }
    }
    return links;
  })(dataLinks)

  // list of source Spreadsheets opened by ids;
  const srcSpreadsheets = sheetIds.reduce(
    (spreadsheets, sheetId)  => {
      try{
        const spreadsheet = SpreadsheetApp.openById(sheetId);
        spreadsheets.push(spreadsheet);
        return spreadsheets;
      } catch(e){
        v>0&& log('When opening sheet with id\n', sheetId, '\ngot: ', e);
        return spreadsheets;
      }
    }, []
  );

  if ( ! srcSpreadsheets.length){
    v>0&& log('No source spreadsheets opened!'); 
    return 2;
  } else {
    v>1&& log(`Spreadsheets opened ${srcSpreadsheets.length}, [${srcSpreadsheets.map(ss => ss.getName())}]`);
  }

  const foundRecords = new Map(); 
  for (const sheet of srcSpreadsheets){
    for (const [dateStr, record] of searchRecords(sheet)){
      foundRecords.set(dateStr, record);
    }
  }

  if (!foundRecords.size)
    v>1&& log(`No records found in spreadsheet ${srcSpreadsheets[0].getName()}`);
  else
    v>2&& log(`Found ${foundRecords.size} day-records in spreadsheets: [${srcSpreadsheets.map(ss => ss.getName())}]`);

  
  // retrieve existing records in raw data sheet
  const existingRecords = getRecords(sheetToImportTo.getRange('A2:F'));
  v>2&& log(`${existingRecords.size} day-records exists in ${sheetToImportTo.getName()}`);

  const dates = datesBetween(fromDate, toDate);
  v>2&& log(`Searching found-records between ${new Date(fromDate).toLocaleDateString('ro-RO')} and ${new Date(toDate).toLocaleDateString('ro-RO')}...`);

  for (const dateStr of dates){
    const foundDateRecords = foundRecords.get(dateStr);
    const existingDateRecords = existingRecords.get(dateStr);
    if (foundDateRecords && existingDateRecords){
      v>1&& log(`Duplicates found on date ${new Date(dateStr).toLocaleDateString('ro-RO')}.`);
      v>2&& log('resolving same-date-key conflicts...');
      const mergedDateRecords = mergeDateRecords(foundDateRecords, existingDateRecords);
      existingRecords.set(dateStr, mergedDateRecords);
    } else if (foundDateRecords){
      existingRecords.set(dateStr, foundDateRecords);
    }
  }

  v>2&& log('updating raw data sheet...')
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
  v>2&& log(`Deleted all 'A2:F' values from sheet ${sheetToImportTo.getName()}!`); 
  // writing new values
  
  v>2&& log(`Writing new values...`);
  rawDataRange.setValues(rawValues);


v>0&& log('Procedure importData END');
} // importData END



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
        v>2&& log(`Record ${Array.from(map1.values())} is duplicate, so skipped`); 
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
function searchRecords(spreadsheet, rowLim=50, colLim=6){
  const records = new Map();

  // measurements
  const messages = new Map();

  // pattern to search against 
  const identifierRe = /=RIGHT\(CELL\("filename",A\d\),LEN\(CELL\("filename",A\d\)\)-FIND\("\]",CELL\("filename",A\d\)\)\)/;

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
v>1&& messages.forEach((vals, mess) => log(mess, vals.length, JSON.stringify(vals)));
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
  
function getRecords(range){
  const rangeValues = range.getValues();
  const records = new Map();
   
  for (const row of rangeValues){
    const record = new Map();
    if (row.filter(v=>v!="").length){
      const date = row[0].toJSON();
      record.set('date', date);
      record.set('ref', row[1]);
      record.set('doc_type', row[2]);
      record.set('descr', row[3]);
      record.set('I_O_type', row[4]);
      record.set('value', row[5]);
      // a record will be retrieved by date
      records.get(date) && records.get(date).push(record)
        || records.set(date, [record]);
    } 
  }
  return records;
}

function updateRawDataSheetNames(rawDataSheets, computedNames){
      rawDataSheets.map(
        (sheet, i) => sheet.setName(computedNames[i])
      )
}

/**
 * Logging function - logs to specified cell
 *      - instantiate with const log = Log(spreadsheetId, sheetIndex, cellPos);
 *      - usage: log("Welcome to log console!");
 *
 * @param {Sheet} spreadsheetId - https://docs.google.com/spreadsheets/d/spreadsheetId/edit#gid=0
 * @param {Number} sheetIndex - numeric index (including 0) of {Sheet} targeted
 * @param {Array} cellPos - tuple array with cell position [x, y] - console output cell
 */
function Log(spreadsheetId, sheetIndex, cellPos){
  const spreadSheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadSheet.getSheets()[sheetIndex];
  // clear console space
  sheet.getRange(...cellPos,8,3).clear();
  const range = sheet.getRange(...cellPos,8,3).merge();
  const cell = range.getCell(1,1);
  cell.setBackground("black");
  cell.setFontColor("white");
  cell.setVerticalAlignment("top");
  
  const logs = [];

  const log = (...args) => {
    logs.push(args.toString());
    cell.setValue('> '+logs.join('\n> '));
    // returns true to permit chaining like: log('something') && another_statement;
    return true;
  }
  return log;
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
    total:{
      label_element:{cell:[15,4],value:"Total la data de {}:",
        style:LABEL_STYLE},
      target_element:{cell:[15,5]}
      
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

} // makeReport END



