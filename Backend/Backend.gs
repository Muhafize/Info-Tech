// This is the main code file for the web app
// This file contains the code for the server side functions
function doGet() {
  return HtmlService.createTemplateFromFile("index.html")
    .evaluate()
    .setTitle("Google Apps Script - Web App Template")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function openApp() {
  let service = ScriptApp.getService();
  if (service.isEnabled()) {
    let ui = SpreadsheetApp.getUi();
    let url = service.getUrl().replace("/dev", "/exec");
    let html = `<script>window.open("${url}");google.script.host.close();<\/script>`;
    let userInterface = HtmlService.createTemplate(html)
      .evaluate()
      .setTitle("Opening ...");
    ui.showSidebar(userInterface);
  } else {
    SpreadsheetApp.getActive().toast(
      "Please deploy the project as web app first.",
      "Message"
    );
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("App").addItem("Open app", "openApp").addToUi();
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}
function getPageUrl(page) {
  let url = ScriptApp.getService().getUrl().replace("/dev", "/exec");
  if (page) {
    url += "?p=" + page;
  }
  return url;
}

function createFile({ data, type, name }, folder) {
  data = Utilities.base64Decode(data);
  const blob = Utilities.newBlob(data, type, name);
  const file = folder.createFile(blob);
  return file;
}

function getFolderByName(name) {
  const id = SpreadsheetApp.getActive().getId();
  const parentFolder = DriveApp.getFileById(id).getParents().next();
  const folders = parentFolder.getFoldersByName(name);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = parentFolder.createFolder(name);
  }
  return folder;
}

// This function is used to save the Schema and app settings to the script properties by reading from the App Settings sheet and Schema sheet
const saveAppPrefs = () => {
  const obj = {};
  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("App Settings");
  const appSettings = settingsSheet
    .getRange(1, 1, settingsSheet.getLastRow(), settingsSheet.getLastColumn())
    .getValues();
  const appSettingsHeaders = appSettings.shift();
  const appSettingsJsonArray = appSettings.map((row) => {
    return row.reduce((obj, value, index) => {
      obj[appSettingsHeaders[index]] = value;
      return obj;
    }, {});
  });
  obj.appSettings = appSettingsJsonArray[0];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    obj.appSettings.SchemaSheet
  );

  const data = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  const headers = data.shift();
  const jsonArray = data.map((row) => {
    return row.reduce((obj, value, index) => {
      obj[headers[index]] = value;
      return obj;
    }, {});
  });

  obj.schema = jsonArray;

  // save to script properties
  PropertiesService.getScriptProperties().setProperty(
    "appPrefs",
    JSON.stringify(obj)
  );
  return jsonArray;
};

// This function is used to get the appPrefs from the script properties
function getAppPrefs() {
  const appPrefs =
    PropertiesService.getScriptProperties().getProperty("appPrefs");
  return JSON.parse(appPrefs);
}

// This function is used to populate the dropdown list of sheets in the App Settings sheet
function getAllSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const sheetNames = sheets.map((sheet) => sheet.getName());

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(sheetNames)
    .build();
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("App Settings");
  sheet.getRange("D2:E2").setDataValidation(rule);
}

// ----------------- ORM Class -------------------
class ORM {
  constructor() {
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      getAppPrefs().appSettings.DataEntrySheet
    );
    this.ID_COL = getAppPrefs().appSettings.IdColumn;
    this.schema = getAppPrefs().schema;
  }

  // Create a new record
  create(data) {
    const id = this.getNextId() || 1;
    data[this.ID_COL] = id;

    const headers = this.sheet
      .getRange(1, 1, 1, this.sheet.getLastColumn())
      .getValues()[0];

    const newRow = [];

    for (const header of headers) {
      newRow.push(data[header] || "");
    }
    this.sheet.appendRow(newRow);
  }

  // Read all records
  readAll() {
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const records = [];

    //Returning data from multi-dimensional array
    for (let i = 1; i < values.length; i++) {
      const record = {};
      for (let j = 0; j < headers.length; j++) {
        record[headers[j]] = values[i][j];
      }
      records.push(record);
    }
    Logger.log(records);
    return records.reverse();
  }

  // Read a specific record by ID
  readById(id) {
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        const record = {};
        for (let j = 0; j < headers.length; j++) {
          record[headers[j]] = values[i][j];
        }
        return record;
      }
    }
    return null;
  }

  // Update a record by ID
  updateById(data) {
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] == data[this.ID_COL]) {
        this.schema.forEach(({ key, type }) => {
          const columnIndex = headers.indexOf(key);
          if (columnIndex !== -1) {
            values[i][columnIndex] = data[key];
            if (type != "formula") {
              this.sheet.getRange(i + 1, columnIndex + 1).setValue(data[key]);
            }
          }
        });
        return true;
      }
    }

    return false;
  }

  // Delete a record by ID
  deleteById(record) {
    const id = record[this.ID_COL];
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] == id) {
        this.sheet.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  }

  // Get the next ID
  getNextId() {
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    let maxId = 0;
    for (let i = 1; i < values.length; i++) {
      const id = values[i][0];
      if (id > maxId) {
        maxId = id;
      }
    }
    return maxId + 1;
  }
}
const getSheetData = (name) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ""), o), {})
  );
  return JSON.stringify(obj);
};

// ----------Callables Methods ------------

// Create a new record
function createRecord(data) {
  const orm = new ORM();
  orm.create(data);
  return readAllRecords();
}

// Read all records
function readAllRecords() {
  const orm = new ORM();
  const allRecords = orm.readAll();
  return JSON.stringify(allRecords);
}

// Read a specific record by ID
function readRecordById(recordId) {
  const orm = new ORM();
  const specificRecord = orm.readById(recordId);
}

// Update a record by ID
function updateRecordById(data) {
  const orm = new ORM();
  const isUpdated = orm.updateById(data);
  return readAllRecords();
}

// Delete a record by ID
function deleteRecord(record) {
  const orm = new ORM();
  const isDeleted = orm.deleteById(record);
  return readAllRecords();
}
