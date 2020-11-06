# Cash money daily reports generator
## Google Sheets apps script 

### Usage:
- Open your Chromium based browser (Chrome)
- Add you Google account (something like "Sign in to Chrome")
- open [Google_Sheets](https://docs.google.com/spreadsheets/u/0/)
- add a new blank sheet
- ...
- go to Tools -> Script Editor (this will open a new tab "https://script.google.com/")
- make sure you see your google email account in top-right of scripts page
- to reveal "appsscript.json" go to View -> Show manifest file
- on the left you will see two files: "Code.js" and "appsscript.json"
- replace content of Code.js with [this_source_code](https://raw.githubusercontent.com/CosminEugenDinu/google-sheets-cash-money-report/master/src/Code.js?token=AIUO72HZW4QZBXJVQ6QDKYK7UU7PQ) and save it (Ctrl+S)
- replace content of appsscript.json with [this_source_code](https://raw.githubusercontent.com/CosminEugenDinu/google-sheets-cash-money-report/master/src/appsscript.json?token=AIUO72GGV2Q3GOEJABET56S7UU5QO) and save it (Ctrl+S)
- go back to your spreadsheet tab, create a button (drawing) and assign it a script ***main***
(go to Insert -> Drawing -> `draw a shape` -> Save and close -> `click on shape` -> Assign a script -> `type` **main** -> OK)
- to run it, click on that button
- a pop-up "Authorization Required" might appear; click Continue -> `chose your account` -> Alert "This app isn't verified" -> Advanced -> Click `Go to your_project_name (unsafe)` -> Allow
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
echo "{\"scriptId\":\"$SCRIPT_ID\"}" > .clasp.json
)
```
- now you have a .clasp.json file with scriptId
- push files to google:
```bash
clasp push
```

