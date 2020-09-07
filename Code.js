

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
  companyName:{
    label:{cell:[2,2], value:"Societatea",
      style:{fontSize:8, fontWeight:"bold", horizontalAlignment:'left'}},
    target:{cell:[2,4]}},
  tax_id:{
    label:{cell:[3,2], value:"CUI",
      style:{fontSize:8, fontWeight:'bold', horizontalAlignment:'left'}},
    target:{cell:[3,4],
      style:{horizontalAlignment:'left'}}},
  reg_num:{
    label:{cell:[4,2], value:"Nr. Reg. Com.:", 
      style:{fontSize:8, fontWeight:'bold', horizontalAlignment:'left'}},
    target:{cell:[4,4]}},
  title:{
    label:{cell:[8,2], value:"REGISTRUL DE CASA", offset:[2,4],
      style:{fontSize:12, fontWeight:'bold', horizontalAlignment:'center', verticalAlignment:'middle'}}},
  document:{
    label:{cell:[11,1], value:"Document", offset:[1, 3],
      style:LABEL_STYLE}},
  explanations:{
    label:{cell:[11,4], value:"EXPLICATII", offset:[2,1],
      style:LABEL_STYLE}},
  input:{
    label:{cell:[11,5], value:"INCASARI", offset:[2,1],
      style:LABEL_STYLE}},
  output:{
    label:{cell:[11,6], value:"PLATI", offset:[2,1],
      style:LABEL_STYLE}},
  date:{
    label:{cell:[12,1], value:"DATA",
      style:LABEL_STYLE}},
  ref:{
    label:{cell:[12,2], value: "NR",
      style:LABEL_STYLE}},
  doc_type:{
    label:{cell:[12,3], value:"TIP",
    style:LABEL_STYLE}},
  previous_balance:{
    label:{cell:[13,4], value:"SOLD LUNA/ZIUA PRECEDENTA"},
    target:{cell:[13,4]},
  total:{
    label:{cell:[15,4],value:"Total la data de {}:"},
  day_balance:{
    label:{cell:[15,5],value:"Sold la data de {}:"},
};



function makeReport() {
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
const [fromDateStr, toDateStr] = [fromDate.toJSON(), toDate.toJSON()];

// data source sheet coresponds with chosen company alias from drop-down
const srcRawDataSheet = repGenSprSheet.getSheetByName(companyAlias+RAWDATA_SHEET_SUFFIX);

//-----------------------------------------------------------------
renderReport(repSprSheet, srcRawDataSheet, fromDateStr, toDateStr);
//-----------------------------------------------------------------


// ------ library -----------------------------------------------------------------


function renderReport(toSpreadsheet, srcRawDataSheet, fromDateStr, toDateStr){

  class Element {

    /**
     * @param {Object} elem - a TEMPLATE property exept those begining with "_"
     */
    constructor(elem){
      const instanceProps = ['cell', 'offset', 'value', 'style'];
      for (const p of instanceProps) this[p] = null;
      if (elem) for (const prop of instanceProps)
        if (elem.hasOwnProperty(prop)) this[prop] = elem[prop];
    }

    setProperty(range, property, value){
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
      if (!this.cell) throw new Error("Cannot render element if no cell");
      const range = sheet.getRange(...this.cell);
      if (this.offset) sheet.getRange(...this.cell, ...this.offset).merge();
      const cell = range.getCell(1,1);
      //cell.setValue(this.value);
      range.setValue(this.value);

      for (const prop in this.style)
        this.setProperty(range, prop, this.style[prop]);

      return range
    }


  }

  
  class DailyReport {
    constructor(dayDate, company, dataValues){
      this.dateStr = dayDate;
      this.formatedDateStr = dayDate.slice(0,10).split('-').reverse().join('/');
      this.company = company;
      this.values = dataValues;
    }
    
    setColumnWidths(sheet, widths){
      widths.map((w, i) => sheet.setColumnWidth(i+1, w));
    }

    render(toSheet, template){
      toSheet.setName(this.formatedDateStr);
      toSheet.getRange(...template._layoutRange).clear();
      // groups of elements 
      const groups = new Map(); 
      for (const entityKey in TEMPLATE){
        if ( entityKey.charAt(0) === '_') continue;
        // elemTypes could be label, target
        const elemTypes = TEMPLATE[entityKey];
        for (const elemType in elemTypes){
          const element = new Element(elemTypes[elemType]);
          if (groups.has(entityKey)){
            groups.get(entityKey).set(elemType, element);
          } else {
            groups.set(entityKey, new Map().set(elemType, element));
          }
        }
      }
      // set dynamic values in target elements
      groups.get('companyName').get('target').value = company.get('name');
      groups.get('tax_id').get('target').value = company.get('tax_id');
      groups.get('reg_num').get('target').value = company.get('reg_num');
      
      for (const [group, elemTypes] of groups){
        for (const [type, element] of elemTypes){	
          element.render(toSheet);
        }
      }

      log(`DailyReport rendered to sheet ${toSheet.getName()}`);
      
    }
    

  }
  
  
  const dataRange = srcRawDataSheet.getRange('A2:F');
  const records = getRecords(dataRange);
  const companyAlias = srcRawDataSheet.getSheetName().replace(RAWDATA_SHEET_SUFFIX, "");
  const company = companies.get(companyAlias);
  const dayTrades = records.get(fromDateStr);
//  const dates = datesBetween(fromDateStr, toDateStr);

  const dayReport = new DailyReport(fromDateStr, company, dayTrades);
  dayReport.render(toSpreadsheet.getSheets()[0], TEMPLATE);

return;
  
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

  const log = str => {
    logs.push(str);
    cell.setValue('> '+logs.join('\n> '));
  }
	return log;
}


}

