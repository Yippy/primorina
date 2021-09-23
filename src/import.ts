// for license and source, visit https://github.com/3096/primorina

function writeLogToSheet(logSheetInfo: ILogSheetInfo) {
  const config = getConfig();
  const reasonMap = getReasonMap(config);
  const settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);

  let authKey: string, serverDivide: ServerDivide;
  try {
    authKey = getParamValueFromUrlQueryString(config.authKeyUrl, API_PARAM_AUTH_KEY);
    serverDivide = getServerDivideFromUrl(config.authKeyUrl);

  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Invalid Auth Key URL: ${err}`, "Error");
    settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("No URL");
    return;
  }

  const endpoint = getApiEndpoint(logSheetInfo, serverDivide);

  const params = getDefaultQueryParams();
  params.set(API_PARAM_AUTH_KEY, authKey);
  params.set(API_PARAM_LANG, config.languageCode);
  params.set(API_PARAM_SIZE, "20");

  const logSheet = SpreadsheetApp.getActive().getSheetByName(logSheetInfo.sheetName);
  const curValues = logSheet.getDataRange().getValues();
  const previousRowCount = curValues.length;

  const LOG_HEADER_ROW = curValues[0];
  const ID_INDEX = LOG_HEADER_ROW.indexOf("id");

  const newRows = [];
  const stopAtId = curValues.length > 1 ? curValues[1][ID_INDEX] : null;

  const addLogsToNewRows = (newLogs: LogEntry[]) => {
    for (const log of newLogs) {
      const newRow = [];
      for (const col of LOG_HEADER_ROW) {
        // handle special cols first
        switch (col) {
          case "date":
            newRow.push(log.time);
            continue;
          case "detail":
            const reasonId = parseInt(log.reason);
            newRow.push(reasonMap.get(reasonId));
            continue;
        }

        if (log.hasOwnProperty(col)) {
          newRow.push(log[col]);
          continue;
        }

        newRow.push("");
      }

      newRows.push(newRow);
    }
  };

  while (true) {
    const response: ApiResponse = requestApiResponse(endpoint, params);
    const entries = response.data.list;
    if (entries.length === 0) {
      // reached the end of logs
      break;
    }

    if (stopAtId) {
      const stopAtIdx = entries.findIndex(entry => entry.id === stopAtId);
      if (stopAtIdx >= 0) {
        addLogsToNewRows(entries.slice(0, stopAtIdx));
        break;
      }
    }

    addLogsToNewRows(entries);
    params.set(API_END_ID, entries[entries.length - 1].id);
  }
  if (newRows.length > 0) {
    newRows.push(...curValues.slice(1));
    logSheet.getRange(2, 1, newRows.length, LOG_HEADER_ROW.length).setValues(newRows);
  }
  const dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
  dashboardSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_dashboard_length']).setValue(newRows.length);
  settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("Found: " + ((newRows.length + 1) - (previousRowCount)));
}

const getPrimogemLog = () => writeLogToSheet(PRIMOGEM_SHEET_INFO);
const getCrystalLog = () => writeLogToSheet(CRYSTAL_SHEET_INFO);
const getResinLog = () => writeLogToSheet(RESIN_SHEET_INFO);


const LOG_RANGES = {
  "Primogem Log": { "range_status": "E26", "range_toggle": "E19", "range_dashboard_length": "C15" },
  "Crystal Log": { "range_status": "E27", "range_toggle": "E20", "range_dashboard_length": "C20" },
  "Resin Log": { "range_status": "E28", "range_toggle": "E21", "range_dashboard_length": "C25" },
};

function importFromAPI() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
  settingsSheet.getRange("E24").setValue(new Date());
  settingsSheet.getRange("E25").setValue("");

  var logName;
  var bannerSheet;
  var bannerSettings;
  // Clear status
  for (var i = 0; i < NAME_OF_LOG_HISTORIES.length; i++) {
    logName = NAME_OF_LOG_HISTORIES[i];
    bannerSettings = LOG_RANGES[logName];
    settingsSheet.getRange(bannerSettings['range_status']).setValue("");
  }
  for (var i = 0; i < NAME_OF_LOG_HISTORIES.length; i++) {
    logName = NAME_OF_LOG_HISTORIES[i];
    bannerSettings = LOG_RANGES[logName];
    var isToggled = settingsSheet.getRange(bannerSettings['range_toggle']).getValue();
    if (isToggled == true) {
      bannerSheet = SpreadsheetApp.getActive().getSheetByName(logName);
      if (bannerSheet) {
        if (logName == SHEET_NAME_CRYSTAL_LOG) {
          getCrystalLog();
        } else if (logName == SHEET_NAME_PRIMOGEM_LOG) {
          getPrimogemLog();
        } else if (logName == SHEET_NAME_RESIN_LOG) {
          getResinLog();
        } else {
          settingsSheet.getRange(bannerSettings['range_status']).setValue("Error log sheet");
        }
      } else {
        settingsSheet.getRange(bannerSettings['range_status']).setValue("Missing sheet");
      }
    } else {
      settingsSheet.getRange(bannerSettings['range_status']).setValue("Skipped");
    }
  }

  settingsSheet.getRange("E25").setValue(new Date());
}
