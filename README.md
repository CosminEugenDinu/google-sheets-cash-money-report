# Cash money daily reports generator
## Google Sheets apps script 
This application has three main functions:
- `importData` - parse old-formatted spreadsheet reports, identify data categories and intercalate them with current working data;
- `cleanRawData` - efficient algorithm for data type checking (with predefined field validators), automatic type conversion, duplicates removal and sorting data records;
- `renderReport` - validate records and creates daily report sheets (using a customizable json mini-template) for archiving or printing.

### Usage:
- Open your Chromium based browser (Chrome)
- Add you Google account (something like "Sign in to Chrome")
- the next link is a template spreadsheet; open it in your browser: [report-generator-template](https://docs.google.com/spreadsheets/d/1MPF0Cbu1mjP36JNDn6GASF73LFX6tEgQi5hZeVD49d4/edit?usp=sharing)
- make you own copy: File -> Make a copy
- in you copy go to Tools -> Script Editor (this will open a new tab "https://script.google.com/")
- make sure you see your google email account in top-right of scripts page
- to reveal "appsscript.json" go to View -> Show manifest file
- on the left you will see two files: "Code.js" and "appsscript.json"
- replace content of Code.js with [this_source_code](https://raw.githubusercontent.com/CosminEugenDinu/google-sheets-cash-money-report/master/src/Code.js) and save it (Ctrl+S)
- replace content of appsscript.json with [this_source_code](https://raw.githubusercontent.com/CosminEugenDinu/google-sheets-cash-money-report/master/src/appsscript.json) and save it (Ctrl+S)
- reports will be sent to a specific spreadsheet, like `Reports`. Let't create it. Go to [Google Sheets](https://docs.google.com/spreadsheets/u/0/), create it, then copy it's id. A spreadsheet ID can be extracted from its URL. For example, the spreadsheet ID in the URL https://docs.google.com/spreadsheets/d/abc1234567/edit#gid=0 is "abc1234567".
- paste this id in spreadsheet `Copy of report-generator-template`, sheet `settings`, column `procedure.variable.value`, row `2`.
- to run the script and generate reports, go to sheet `Interface` and click on *that* button (make sure you have selected a reasonable date range)
- a pop-up "Authorization Required" might appear; click Continue -> `chose your account` -> Alert "This app isn't verified" -> Advanced -> Click `Go to makeReport (unsafe)` -> Allow
- if all worked (no errors appeared) you should see your reports in `Reports` spreadsheet
- done!

## Development - debugging:
- requirements: GNU-Linux-distro, [nodejs](https://nodejs.org/), [npm](https://www.npmjs.com/get-npm/), [clasp](https://github.com/google/clasp)
```bash
(
clasp login 
# you will see `Default credentials saved to: ~/.clasprc.json (/home/user/.clasprc.json).` 
git clone https://github.com/CosminEugenDinu/google-sheets-cash-money-report.git
cd google-sheets-cash-money-report/src
)
```
- go to https://script.google.com/home and select the project 
- got to File -> Project properties -> Script ID -> you will copy this and paste it after running next script
```bash
(
read -p "Paste here your Script ID:" SCRIPT_ID
echo "{\"scriptId\":\"$SCRIPT_ID\"}" > .clasp.json && echo "File .clasp.json created!"
)
```
- now you have a .clasp.json file with scriptId
- in order to push files to google, go to https://script.google.com/home/usersettings -> Google Apps Script API -> ON
```bash
# push local changes to google scripts
clasp push
```
- testing:
```bash
cd test
npm install
node nodejs_test.js
```

## Description
- purpose
- how to use
- how it works

![Interface sheet](/docs/images/Interface.png)
![Reports](/docs/images/Reports.png)

