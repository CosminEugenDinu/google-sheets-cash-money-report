
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

const LABEL_STYLE = {fontSize:8, horizontalAlignment:'center', verticalAlignment:'middle', background:'lightgray', borders:[true, true, true, true, true, true]};

const TEMPLATE = {
  _layoutRange:[1,1,50,6],
  _columnWidths:[75,80,20,250,80,80],
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
    data:{
      target_element:{cell:[14,1]},
    },
    ref:{
      target_element:{cell:[14,2]},
    },
    doc_type:{
      target_element:{cell:[14,3]},
    }},
  total:{
    label_element:{cell:[15,4],value:"Total la data de {}:",
      style:LABEL_STYLE}},
  day_balance:{
    label_element:{cell:[16,4],value:"Sold la data de {}:",
      style:LABEL_STYLE}},
  body:{
    frame_element:{cell:[13,1], extent:[4,6],
      style:{borders:[null, true, true, true, false, false]}},
  },
  
};


// instantiate log function
const log = Log(REPORT_GENERATOR_SPREADSHEET_ID, 0, [10,5]);
log("makeReport procedure begin...");
  
const repGenSprSheet = SpreadsheetApp.openById(REPORT_GENERATOR_SPREADSHEET_ID);
const repSprSheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
 
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

// user choose company alias from drop-down in interface
const [[companyAlias]] = interface.getSheetValues(8,2,1,1);
// user selects in interface
const [[fromDate, toDate]] = interface.getSheetValues(8,3,1,2);

// data source sheet corresponds with chosen company alias from drop-down
const srcRawDataSheet = repGenSprSheet.getSheetByName(companyAlias+RAWDATA_SHEET_SUFFIX);

const company = companies.get(companyAlias);
//const dataRange = srcRawDataSheet.getRange('A2:F');
//const records = getRecords(dataRange);
//const companyAlias = srcRawDataSheet.getSheetName().replace(RAWDATA_SHEET_SUFFIX, "");
//const company = companies.get(companyAlias);
//const dayTrades = records.get(fromDate.toJSON());
//const dates = datesBetween(fromDate, toDate);

const dataRecords = new Map();
const targetSpreadsheet = repSprSheet;

//-----------------------------------------------------------------
renderReport(fromDate, toDate, company, dataRecords, targetSpreadsheet);
//-----------------------------------------------------------------









// ------ library -----------------------------------------------------------------

function renderReport(fromDate, toDate, company, dataRecords){

/**
 * Class Element - is a piece of sheet... (cell, range)
 *
 * Depending on type of parentKey (e.g. 'target_element', 'label_element', 'frame_element'),
 * assigns specific properties (e.g. only element 'frame_element' has property 'extent')
 * When render method is called, that element produces effect on target sheet,
 * like setting a value in a cell or changing background color.
 */
class Element {

  /**
   * @param {Object} elem - a TEMPLATE property except those beginning with "_"
   * @param {string} parentKey - key in {Map} tree where elem is stored
   */
  constructor(parentKey, elem){

    const typesProps = Element._typesProps

    if (!this.supportedTypes.includes(parentKey))
      throw new TypeError(`${parentKey} is not a valid element type`+
        `supported types are: ${this.supportedTypes}`);
    this._type = parentKey;

    for (const prop of Element._typesProps.get(parentKey))
      this[prop] = null;
    
    for (const p in elem){
      if (typesProps.get(this._type).includes(p)){
        this[p] = elem[p];
      } else {
        throw new TypeError(`Property ${p} is not a supported by element type ${this._type}`);
        }
    }

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

    return range
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
 */
class DailyReport {

  constructor(date, company, dataValues){
    this.date = date;
    this.company = company;
    this.values = dataValues;
  }
  
  setColumnWidths(sheet, widths){
    widths.map((w, i) => sheet.setColumnWidth(i+1, w));
  }
  setRowsHeight(sheet, numRows, height){
    sheet.setRowHeights(1, numRows, height);
  }

  /**
   * Converts tuple array like [1, 2] into {string} key '1:2'
   */
  static keyFromCell(x, y){
    const key = `${x}:${y}`;
    return key;
  }

  /**
   * Converts {string} key (e.g. '1:2') into tuple array (e.g. [1, 2])
   */
  static cellFromKey(key){
    const tuple = key.split(':').map(letter=>+letter);
    return tuple // cell 
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
        const [x, y] = obj[key].cell;
        leaves.set(DailyReport.keyFromCell(x, y), element); 
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
    
    toSheet.setName(this.date.toLocaleDateString());
    this.setColumnWidths(toSheet, template._columnWidths);
    const numRows = template._layoutRange[2];
    this.setRowsHeight(toSheet, numRows, template._rowHeight);
    toSheet.getRange(...template._layoutRange).clear();

    const leafKeys = Element.getSupportedTypes();
    // {Map} tree - having {Element} leaves
    // {Map} elements - having key=DailyReport.keyFromCell(x,y), and value is {Element} leaf 
    const [tree, elements] = DailyReport.objToMap(template, leafKeys);
    // populate headers (general info displayed on top of report sheet)
    tree.get('companyName').get('target_element').value = company.get('name');
    tree.get('tax_id').get('target_element').value = company.get('tax_id');
    tree.get('reg_num').get('target_element').value = company.get('reg_num');

    // change label according to date (if date is 1st or not)
    let label = tree.get('previous_balance').get('label_element');
    if (this.date.getDate() === 1)
      label.value = label.value.replace(/\/ziua/i, '');
    else 
      label.value = label.value.replace(/luna\//i, '');
    // replace '{}' with date in corresponding labels
    label = tree.get('total').get('label_element');
    label.value = replaceCurly(label.value, this.date.toLocaleDateString()); 
    label = tree.get('day_balance').get('label_element');
    label.value = replaceCurly(label.value, this.date.toLocaleDateString()); 
    
    for (const [key, element] of elements){
      element.render(toSheet);
      }
  }
}



class Report{
  constructor(fromDate, toDate, company, dataRecords){
    this.fromDate = fromDate;
    this.toDate = toDate;
    this.company = company;
    this.dataRecords = dataRecords;
  }
  
  render(targetSpreadsheet){
    /* for every date between fromDate and toDate:
     *   collect dataRecords and group by date in a {Map},
     *   generate an instance of {DailyReport},
     *   create a new {Sheet} instance in {Spreadsheet} and name it with date,
     *   render every dayReport to sheet according with date,
     *   and DONE
     */

    const dayTrades = dataRecords.get(fromDate.toJSON());
    //const dates = datesBetween(fromDate, toDate);

    const dayReport = new DailyReport(fromDate, company, dayTrades);
    dayReport.render(targetSpreadsheet.getSheets()[0], TEMPLATE);
    return;
  }


}

//----------------------------------------------------------------
//-------------- render all reports ------------------------------
const report = new Report(fromDate, toDate, company, dataRecords)
report.render(targetSpreadsheet);
//================================================================


} // renderReport END













// ----------Global functions (in makeReport scope)---------------

function getCompanies(sheet, records=10, fields=4){
  const companies = new Map();
  let company_id = 0
  for (const row of sheet.getSheetValues(2,1,records,fields)){
    if (row.filter(val => val === "").length){
      company_id ++;
      continue;
      }
    const company = new Map();
    company.set('id', company_id);
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
  
  let i = 0;
  for (const row of rangeValues){
    const record = new Map();
    if (row.filter(v=>v!="").length){
      record.set('id', i);
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
    i++
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

} // makeReport END


