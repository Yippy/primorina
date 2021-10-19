// for license and source, visit https://github.com/3096/primorina

function writeLogToSheet(logSheetInfo: ILogSheetInfo) {
  const config = getConfig();
  const reasonMap = getReasonMap(config);
  const settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
  const isHoYoLAB = logSheetInfo.sheetName == SHEET_NAME_MORA_LOG;
  let authKey: string, serverDivide: ServerDivide;
  if (isHoYoLAB) {
    const serverSetting = settingsSheet.getRange("B3").getValue();
    if (serverSetting == "China") {
      serverDivide = "cn";
    } else {
      serverDivide = "os";
    }
  } else {
    try {
      authKey = getParamValueFromUrlQueryString(config.authKeyUrl, API_PARAM_AUTH_KEY);
      serverDivide = getServerDivideFromUrl(config.authKeyUrl);
    } catch (err) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Invalid Auth Key URL: ${err}`, "Error");
      settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("No URL");
      return;
    }
  }
  const endpoint = getApiEndpoint(logSheetInfo, serverDivide);

  const params;
  var currentPage;
  var selectedMonth;
  if (isHoYoLAB) {
    params = getDefaultQueryParamsForHoYoLab();
    params.set(API_PARAM_REGION, config.regionCode);
    params.set(API_PARAM_LANG, config.languageCode);
    currentPage = 0;
    params.set(API_PARAM_CURRENT_PAGE, currentPage);
  } else {
    params = getDefaultQueryParams();
    params.set(API_PARAM_AUTH_KEY, authKey);
    params.set(API_PARAM_LANG, config.languageCode);
    params.set(API_PARAM_SIZE, "20");
  }
  const logSheet = SpreadsheetApp.getActive().getSheetByName(logSheetInfo.sheetName);
  const curValues = logSheet.getDataRange().getValues();
  const previousLogCount = curValues.length - 1;

  const LOG_HEADER_ROW = curValues[0];
  const ID_INDEX = LOG_HEADER_ROW.indexOf("id");
  const TIME_INDEX = LOG_HEADER_ROW.indexOf("time");

  const stopAtId: string = previousLogCount ? curValues[1][ID_INDEX] : null;
  const lastLogDate: number = previousLogCount ? Date.parse(curValues[1][TIME_INDEX]) : null;

  const newRows = [];
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

  goingThroughApiResponses:
  while (true) {
    var entries;
    if (isHoYoLAB) {
      currentPage++;
      params.set(API_PARAM_CURRENT_PAGE, currentPage);
      try {
        const responseHoYo: ApiResponseHoYo = requestApiResponseHoYo(endpoint, params);
        entries = responseHoYo.data.list;
      } catch (err) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Invalid HoYoLAB ltoken or UID: ${err}`, "Error");
        settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("Invalid HoYoLAB ltoken or UID");
        return;
      }
    } else {
      const response: ApiResponse = requestApiResponse(endpoint, params);
      entries = response.data.list;
    }

    if (entries.length === 0) {
      // reached the end of logs
      if (previousLogCount) {
        const userConfirm = SpreadsheetApp.getUi().alert(
          "Warning",
          `${logSheetInfo.sheetName} import could not match the last recorded entry; still import?\n`
          + "(If you've recently imported to this log, something is likely wrong.)",
          SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
        if (userConfirm !== SpreadsheetApp.getUi().Button.OK) {
          settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("Cancelled");
          return;
        }
      }
      if (isHoYoLAB) {
        if (selectedMonth == null) {
          selectedMonth = responseHoYo.data.data_month;
        } else {
          selectedMonth--;
        }
        if (responseHoYo.data.optional_month.indexOf(selectedMonth) > -1) {
          currentPage = 0;
          params.set(API_PARAM_MONTH, selectedMonth);
        } else {
          break;
        }
      } else {
        break;
      }
    }

    // check response agaist last log
    if (previousLogCount) {
      for (const [index, entry] of entries.entries()) {
        if (isHoYoLAB) {
          if (Date.parse(entry.time) === lastLogDate) {
            // found last time
            addLogsToNewRows(entries.slice(0, index));
            break goingThroughApiResponses;
          }
        } else {
          if (entry.id === stopAtId) {
            // found last id
            addLogsToNewRows(entries.slice(0, index));
            break goingThroughApiResponses;
          }
        }

        if (Date.parse(entry.time) < lastLogDate) {
          // found unexpected datetime
          const errMsg = "imported date reached past last entry, check if account correct?\n"
            + `cur entry ${JSON.stringify(entry)}, last entry ${JSON.stringify(curValues[1])}`;
          settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue(errMsg);
          throw Error(errMsg);
        }
      }
    }

    addLogsToNewRows(entries);
    if (!isHoYoLAB) {
      params.set(API_END_ID, entries[entries.length - 1].id);
    }
  }

  const finalRows = newRows.concat(curValues.slice(1));
  if (newRows.length > 0) {
    logSheet.getRange(2, 1, finalRows.length, LOG_HEADER_ROW.length).setValues(finalRows);
  }

  const dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
  dashboardSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_dashboard_length']).setValue(finalRows.length);
  settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("Found: " + (finalRows.length - previousLogCount));
}

const getPrimogemLog = () => writeLogToSheet(PRIMOGEM_SHEET_INFO);
const getCrystalLog = () => writeLogToSheet(CRYSTAL_SHEET_INFO);
const getResinLog = () => writeLogToSheet(RESIN_SHEET_INFO);
const getMoraLog = () => writeLogToSheet(MORA_SHEET_INFO);

const LOG_RANGES = {
  "Primogem Log": { "range_status": "E26", "range_toggle": "E19", "range_dashboard_length": "C15" },
  "Crystal Log": { "range_status": "E27", "range_toggle": "E20", "range_dashboard_length": "C20" },
  "Resin Log": { "range_status": "E28", "range_toggle": "E21", "range_dashboard_length": "C25" },
  "Mora Log": { "range_status": "E39", "range_toggle": "E35", "range_dashboard_length": "C30" },
};

const ltokenInput: string;
const ltuidInput: string;

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

function importFromHoYoLAB() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
  settingsSheet.getRange("E37").setValue(new Date());
  settingsSheet.getRange("E38").setValue("");
  var logName;
  var bannerSheet;
  var bannerSettings;
  // Clear status
  for (var i = 0; i < NAME_OF_LOG_HISTORIES_HOYOLAB.length; i++) {
    logName = NAME_OF_LOG_HISTORIES_HOYOLAB[i];
    bannerSettings = LOG_RANGES[logName];
    settingsSheet.getRange(bannerSettings['range_status']).setValue("");
  }
  for (var i = 0; i < NAME_OF_LOG_HISTORIES_HOYOLAB.length; i++) {
    logName = NAME_OF_LOG_HISTORIES_HOYOLAB[i];
    bannerSettings = LOG_RANGES[logName];
    var isToggled = settingsSheet.getRange(bannerSettings['range_toggle']).getValue();
    if (isToggled == true) {
      bannerSheet = SpreadsheetApp.getActive().getSheetByName(logName);
      if (bannerSheet) {
        if (logName == SHEET_NAME_MORA_LOG) {
          ltuidInput = settingsSheet.getRange("D33").getValue();
          if (ltuidInput.length == 0) {
            const result = displayUserPrompt("Auto Import with HoYoLab",`Enter HoYoLAB UID to proceed\n.`);
            var button = result.getSelectedButton();
            if (button == SpreadsheetApp.getUi().Button.OK) {
              ltuidInput = result.getResponseText();
              if (isHoYoIdCorrect(ltuidInput)) {
                settingsSheet.getRange("D33").setValue(ltuidInput);
                getMoraLog();
              } else {
                settingsSheet.getRange(bannerSettings['range_status']).setValue("HoYoLAB is invalid");
              }
            } else {
              settingsSheet.getRange(bannerSettings['range_status']).setValue("User cancelled process");
            }
          } else {
            getMoraLog();
          }
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
  settingsSheet.getRange("E38").setValue(new Date());
}

function isHoYoIdCorrect(userInput: string) {
  const isValid = false;
  if (userInput.match(/^[0-9]+$/) != null) {
    isValid = true;
  }
  return isValid;
}

function displayUserPrompt(titlePrompt: string, messagePrompt: string) {
  const ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    titlePrompt,
    messagePrompt,
  SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  return result;
}