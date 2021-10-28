// for license and source, visit https://github.com/3096/primorina

function getRowProperties(header: string[], row: string[] | null, props: string[]) {
  if (!row) return [...Array(props.length).fill(null)];
  return props.map(prop => {
    const idx = header.indexOf(prop);
    return idx >= 0 ? row[idx] : null;
  });
}

interface SpecialColProcs {
  [specialProp: string]: (entry: any) => string;
}

function addEntriesToRows(
  entries: any[], headerRow: string[], rows: string[][], lastLogTime: number | null,
  stoppingMatch: (entry: any) => boolean, specialColProcs: SpecialColProcs = null)
  : boolean  // returns if stopping match was hit
{
  for (const entry of entries) {
    const newRow = [];
    for (const col of headerRow) {
      if (lastLogTime) {
        // check time
        if (Date.parse(entry.time) < lastLogTime) {
          throw Error("imported date reached past last entry, check if account/server correct?\n"
            + `cur entry ${JSON.stringify(entry)}, last entry time: ${(new Date(lastLogTime)).toString()}`);
        }

        // check for stopping early
        if (stoppingMatch(entry)) return true;
      }

      if (specialColProcs && specialColProcs.hasOwnProperty(col)) {
        newRow.push(specialColProcs[col](entry));
        continue;
      }

      if (entry.hasOwnProperty(col)) {
        newRow.push(entry[col]);
        continue;
      }

      newRow.push("");
    }

    rows.push(newRow);
  }

  return false;
}

function warnLogsNotMatched(logSheetInfo: ILogSheetInfo, settingsSheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const userConfirm = SpreadsheetApp.getUi().alert(
    "Warning",
    `${logSheetInfo.sheetName} import could not match the last recorded entry; still import?\n`
    + "(If you've recently imported to this log, the account is likely wrong.)",
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (userConfirm !== SpreadsheetApp.getUi().Button.OK) {
    settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("Cancelled");
    return false;
  }
  return true;
}

function writeImServiceLogToSheet(logSheetInfo: ILogSheetInfo) {
  const config = getConfig();
  const settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
  const logSheet = SpreadsheetApp.getActive().getSheetByName(logSheetInfo.sheetName);
  const curValues = logSheet.getDataRange().getValues();
  const logHeaderRow = curValues[0];
  const previousLogCount = curValues.length - 1;

  const reasonMap = getReasonMap(config);
  let authKey: string, serverDivide: ServerDivide;
  try {
    authKey = getParamValueFromUrlQueryString(config.authKeyUrl, API_PARAM_AUTH_KEY);
    serverDivide = getServerDivideFromUrl(config.authKeyUrl);
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Invalid Auth Key URL: ${err}`, "Error");
    settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("No URL");
    return;
  }

  let lastLogId: string = null;
  let lastLogDate: number = null;
  if (previousLogCount) {
    const [lastLogIdStr, lastLogDateStr] = getRowProperties(logHeaderRow, curValues[1], ["id", "time"]);
    lastLogId = lastLogIdStr;
    lastLogDate = Date.parse(lastLogDateStr);
  }

  const params = getImServiceDefaultQueryParams();
  params.set(API_PARAM_AUTH_KEY, authKey);
  params.set(API_PARAM_LANG, config.languageCode);
  params.set(API_PARAM_SIZE, "20");

  const endpoint = getApiEndpoint(logSheetInfo, serverDivide);
  const newRows = [];
  let matchedLastLog = false;

  while (true) {
    const response: ImServiceApiResponse = requestApiResponse(endpoint, params);
    const entries = response.data.list;

    if (entries.length === 0) {
      // reached the end of logs
      break;
    }

    const stoppingMatched = addEntriesToRows(entries, logHeaderRow, newRows, lastLogDate,
      (entry: ImServiceLogEntry) => entry.id === lastLogId,
      {
        "detail": (entry: ImServiceLogEntry) => reasonMap.get(parseInt(entry.reason))
      }
    );
    if (stoppingMatched) {
      matchedLastLog = true;
      break;
    }

    params.set(API_END_ID, entries[entries.length - 1].id);
  }

  if (previousLogCount && !matchedLastLog && !warnLogsNotMatched(logSheetInfo, settingsSheet)) {
    return;
  }

  const finalRows = newRows.concat(curValues.slice(1));
  if (newRows.length > 0) {
    logSheet.getRange(2, 1, finalRows.length, logHeaderRow.length).setValues(finalRows);
  }

  const dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
  dashboardSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_dashboard_length']).setValue(finalRows.length);
  settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("Found: " + newRows.length);
}

function writeLedgerLogToSheet(logSheetInfo: ILogSheetInfo) {
  const COMPARING_PROPS = ["time", "num", "action_id"];
  const isSameRowValue = (row0: string[], row1: string[]) => {
    const props0 = getRowProperties(logHeaderRow, row0, COMPARING_PROPS);
    const props1 = getRowProperties(logHeaderRow, row1, COMPARING_PROPS);
    return props0.every((value, idx) => value === props1[idx]);
  };

  const config = getConfig();
  const settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
  const logSheet = SpreadsheetApp.getActive().getSheetByName(logSheetInfo.sheetName);

  let curValues = logSheet.getDataRange().getValues();
  const logHeaderRow = curValues[0];
  const previousLogCount = curValues.length - 1;

  const cookies: Cookies = { ltoken: ltokenInput, ltuid: ltuidInput };
  const serverDivide = REGION_INFO[config.regionCode].serverDivide;

  const params = getLedgerDefaultQueryParams();
  params.set(API_PARAM_REGION, config.regionCode);
  params.set(API_PARAM_LANG, config.languageCode);

  const endpoint = getApiEndpoint(logSheetInfo, serverDivide);
  const curYear = getServerTimeAsUtcNow(config.regionCode).getUTCFullYear();
  const months = (requestApiResponse(endpoint, params, cookies) as LedgerApiResponse).data.optional_month;
  if (months.length === 0) {
    throw Error(`account has no history for ${logSheetInfo.sheetName}`);
  }
  const monthsWithYear =
    months.map(month => `${month <= months[months.length - 1] ? curYear : curYear - 1}-${month}`);

  let hasPreviousLogInRange = true;

  // trim previous log last value
  let trimToRowIdx = 1;
  if (trimToRowIdx < curValues.length) {
    const isInLogRange = (row: string[]) => {
      const [timeStr] = getRowProperties(logHeaderRow, row, ["time"]);
      const apiTimeAsUtc = getApiTimeAsServerTimeAsUtc(timeStr);
      // having getMonth start at Jan = 0 is the stupidest thing ive ever seen
      return monthsWithYear.includes(`${apiTimeAsUtc.getUTCFullYear()}-${apiTimeAsUtc.getUTCMonth() + 1}`);
    }

    if (!isInLogRange(curValues[trimToRowIdx])) {
      hasPreviousLogInRange = false;

    } else {
      do {
        trimToRowIdx++;
        if (trimToRowIdx >= curValues.length || !isInLogRange(curValues[trimToRowIdx])) {
          hasPreviousLogInRange = false;
          break;
        }

      } while (isSameRowValue(curValues[trimToRowIdx], curValues[trimToRowIdx - 1]))

      curValues = [logHeaderRow, ...curValues.slice(trimToRowIdx)];
    }
  } else {
    hasPreviousLogInRange = false;
  }

  let lastImportedLogTime: number = null;
  let lastImportedLogNum: number = null;
  let lastImportedLogAction: number = null;
  if (hasPreviousLogInRange) {
    const [lastImportedLogTimeStr, lastImportedLogNumStr, lastImportedLogActionStr]
      = getRowProperties(logHeaderRow, curValues[1], ["time", "num", "action_id"]);
    lastImportedLogTime = Date.parse(lastImportedLogTimeStr);
    lastImportedLogNum = parseInt(lastImportedLogNumStr);
    lastImportedLogAction = parseInt(lastImportedLogActionStr);
  }

  const newRows = [];
  let matchedLastLog = false;

  goingThroughAllMonths:
  for (const curMonth of months.reverse()) {
    let currentPage = 1;
    while (true) {
      params.set(API_PARAM_CURRENT_PAGE, currentPage.toString());
      params.set(API_PARAM_MONTH, curMonth.toString());
      const response: LedgerApiResponse = requestApiResponse(endpoint, params, cookies);
      const entries = response.data.list;

      if (entries.length === 0) {
        // reached the end of current month logs
        break;
      }

      const stoppingMatched = addEntriesToRows(entries, logHeaderRow, newRows, lastImportedLogTime,
        (entry: LedgerLogEntry) => {
          const { time, num, action_id } = entry;
          return Date.parse(time) === lastImportedLogTime
            && num === lastImportedLogNum
            && action_id === lastImportedLogAction;
        }
      );
      if (stoppingMatched) {
        matchedLastLog = true;
        break goingThroughAllMonths;
      }

      currentPage++;
    }
  }

  if (hasPreviousLogInRange && !matchedLastLog && !warnLogsNotMatched(logSheetInfo, settingsSheet)) {
    return;
  }

  const finalRows = newRows.concat(curValues.slice(1));
  if (newRows.length > 0) {
    logSheet.getRange(2, 1, finalRows.length, logHeaderRow.length).setValues(finalRows);
  }

  const dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
  dashboardSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_dashboard_length']).setValue(finalRows.length);
  settingsSheet.getRange(LOG_RANGES[logSheetInfo.sheetName]['range_status']).setValue("Found: " + (finalRows.length - previousLogCount));
}

const getPrimogemLog = () => writeImServiceLogToSheet(PRIMOGEM_SHEET_INFO);
const getCrystalLog = () => writeImServiceLogToSheet(CRYSTAL_SHEET_INFO);
const getResinLog = () => writeImServiceLogToSheet(RESIN_SHEET_INFO);
const getMoraLog = () => writeLedgerLogToSheet(MORA_SHEET_INFO);

const LOG_RANGES = {
  "Primogem Log": { "range_status": "E26", "range_toggle": "E19", "range_dashboard_length": "C15" },
  "Crystal Log": { "range_status": "E27", "range_toggle": "E20", "range_dashboard_length": "C20" },
  "Resin Log": { "range_status": "E28", "range_toggle": "E21", "range_dashboard_length": "C25" },
  "Mora Log": { "range_status": "E39", "range_toggle": "E35", "range_dashboard_length": "C30" },
};

const ltokenInput: string;
const ltuidInput: string;

function importFromAPI() {
  var settingsSheet = getSettingsSheet();
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
  var settingsSheet = getSettingsSheet();
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
            const result = displayUserPrompt("Auto Import with HoYoLab", `Enter HoYoLAB UID to proceed\n.`);
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
  let isValid = false;
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
