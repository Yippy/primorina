// for license and source, visit https://github.com/3096/primorina

interface SpecialColProcs {
  [specialProp: string]: (entry: any) => string;
}

const sheetTimeToStrProc: SpecialColProcs = {
  'time': (entry: any) => sheetDateToApiTimeStr(new Date(entry)),
};

function getRowProperties(header: string[], row: string[] | null, props: string[], specialColProcs: SpecialColProcs = null) {
  if (!row) throw Error("row is null");
  return props.map(prop => {
    const idx = header.indexOf(prop);
    if (idx == -1) {
      throw Error(`${prop} not found in header`);
    }
    if (specialColProcs && specialColProcs.hasOwnProperty(prop)) {
      return specialColProcs[prop](row[idx]);
    }
    return row[idx];
  });
}

function serverDivideSpecificOp(serverDivide: ServerDivide, ops: { [serverDivide in ServerDivide]: () => any }) {
  return ops[serverDivide]();
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

function findReasonId(reasonDetailName :string, entry: ImServiceLogEntry, reasonMap) {
  let findKey = entry[reasonDetailName];
  let foundIdFromMap = 0;

  // Find mapping first from document, very useful in overriding miHoYo random reason change
  if (localReasonMap == null) {
    getPopulateReasonMap();
  }
  foundIdFromMap = localReasonMap[findKey];

  if (foundIdFromMap == null) {
    // Find from miHoYo
    foundIdFromMap = reasonMap.get(findKey);
  }
  if (foundIdFromMap == null) {
    // Unable to find within document or miHoYo
    foundIdFromMap = 0;
  }
  return foundIdFromMap;
}

var localReasonMap;

function getPopulateReasonMap() {
  let reasonMapSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_REASON_MAP);
  localReasonMap = [];
  let reasonMapData;
  if (reasonMapSheet) {
    reasonMapData = reasonMapSheet.getDataRange().getValues();
  } else {
    reasonMapSheet = loadReasonMapSheet();
    if (reasonMapSheet) {
      reasonMapData = reasonMapSheet.getDataRange().getValues();
    }
  }
  if (reasonMapData) {
    reasonMapData.forEach(function (row, index) {
      if (index > 0) {
        localReasonMap[row[1]] = row[0];
      }
    });
  }
  return reasonMapSheet;
}

function writeImServiceLogToSheet(logSheetInfo: ILogSheetInfo) {
  const config = getConfig();
  const settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
  const logSheet = SpreadsheetApp.getActive().getSheetByName(logSheetInfo.sheetName);
  let curValues = logSheet.getDataRange().getValues();
  let logHeaderRow = curValues[0];
  try {
    // Test if headers are correct before proceeding
    let checkColumns = [logSheetInfo.header.id, logSheetInfo.header.datetime, logSheetInfo.header.ReasonId, logSheetInfo.header.reasonDetail];

    if (logSheetInfo.sheetName == SHEET_NAME_ARTIFACT_LOG || logSheetInfo.sheetName == SHEET_NAME_WEAPON_LOG) {
      checkColumns.push(logSheetInfo.header.itemRarity, logSheetInfo.header.itemLevel,logSheetInfo.header.itemRarity);
    }
    getRowProperties(logHeaderRow, curValues[0], checkColumns);
  } catch (err) {
    addFormulaByLogName(logSheetInfo.sheetName);
    // Reload only the header row
    var logSourceNumberOfColumnWithFormulas = logSheet.getLastColumn();
    logHeaderRow = logSheet.getRange(1, 1, 1, logSourceNumberOfColumnWithFormulas).getValues()[0];
  };

  let previousLogCount = curValues.length - 1;

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
    const [lastLogIdStr, lastLogDateStr] = getRowProperties(logHeaderRow, curValues[1], [logSheetInfo.header.id, logSheetInfo.header.datetime]);
    lastLogId = lastLogIdStr;
    lastLogDate = Date.parse(lastLogDateStr);
  }

  const params = getImServiceDefaultQueryParams(authKey, config.languageCode);

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
        [logSheetInfo.header.total]: (entry: ImServiceLogEntry) => parseInt(entry.add_num), // Conversion needed due to GetCrystalLog outputting string, example '+300'
        [logSheetInfo.header.reasonId]: (entry: ImServiceLogEntry) => findReasonId(logSheetInfo.header.reasonDetail, entry, reasonMap)// There is no reason id, must cross match reason details
      }
    );
    if (stoppingMatched) {
      matchedLastLog = true;
      break;
    }

    params.end_id = entries[entries.length - 1].id;
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

// if import takes too long, cache results before Google kills the run
const importStartTime = Date.now();

function writeLedgerLogToSheet(logSheetInfo: ILogSheetInfo) {
  const COMPARING_PROPS = [logSheetInfo.header.datetime, logSheetInfo.header.total, logSheetInfo.header.reasonId];

  const logSheet = SpreadsheetApp.getActive().getSheetByName(logSheetInfo.sheetName);

  let curValues = logSheet.getDataRange().getValues();
  let logHeaderRow = curValues[0];

  try {
    // Test if headers are correct before proceeding
    let checkColumns = [logSheetInfo.header.id, logSheetInfo.header.datetime, logSheetInfo.header.reasonId, logSheetInfo.header.reasonDetail];
    getRowProperties(logHeaderRow, curValues[0], checkColumns);
  } catch (err) {
    addFormulaByLogName(logSheetInfo.sheetName);
    // Reload only the header row
    var logSourceNumberOfColumnWithFormulas = logSheet.getLastColumn();
    logHeaderRow = logSheet.getRange(1, 1, 1, logSourceNumberOfColumnWithFormulas).getValues()[0];
  };

  const isSameRowValue = (row0: string[], row1: string[], row0Proc = null, row1Proc = null) => {
    const props0 = getRowProperties(logHeaderRow, row0, COMPARING_PROPS, row0Proc);
    const props1 = getRowProperties(logHeaderRow, row1, COMPARING_PROPS, row1Proc);
    return props0.every((value, idx) => value === props1[idx]);
  };

  const config = getConfig();
  const settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);

  const previousLogCount = curValues.length - 1;

  const serverDivide = REGION_INFO[config.regionCode].serverDivide;

  let cookies: Cookies;
  serverDivideSpecificOp(serverDivide, {
    cn: () => { (cookies as LedgerCookieCn) = { cookie_token: ltokenInput, account_id: ltuidInput }; },
    os: () => { (cookies as LedgerCookieOs) = { ltoken: ltokenInput, ltuid: ltuidInput }; }
  });

  const curYear = getServerTimeAsUtcNow(config.regionCode).getUTCFullYear();
  const curMonth = getServerTimeAsUtcNow(config.regionCode).getUTCMonth() + 1;

  const params = LEDGER_GET_DEFAULT_QUERY_PARAMS[serverDivide](config.regionCode);
  params.month = curMonth.toString();
  if (serverDivide === "os") {
    params.lang = config.languageCode;
  }

  const endpoint = getApiEndpoint(logSheetInfo, serverDivide);
  const apiResponseMonthsAvailable
    = (requestApiResponse(endpoint, params, cookies) as LedgerApiResponse).data.optional_month;
  if (apiResponseMonthsAvailable.length === 0) {
    throw Error(`account has no history for ${logSheetInfo.sheetName}`);
  }
  const apiResponseMonthsAvailableWithYear = apiResponseMonthsAvailable.map(month =>
    `${month <= apiResponseMonthsAvailable[apiResponseMonthsAvailable.length - 1] ? curYear : curYear - 1}-${padMonth(month)}`
  );

  let hasPreviousLogInRange = true;

  // trim previous log last value
  let trimToRowIdx = 1;
  if (trimToRowIdx < curValues.length) {
    const rowIsInAvailableLogRange = (row: string[]) => {
      const [timeStr] = getRowProperties(logHeaderRow, row, [logSheetInfo.header.datetime]);
      const apiTimeAsUtc = getApiTimeAsServerTimeAsUtc(timeStr);
      // having getMonth start at Jan = 0 is the stupidest thing ive ever seen
      return apiResponseMonthsAvailableWithYear.includes(
        `${apiTimeAsUtc.getUTCFullYear()}-${padMonth(apiTimeAsUtc.getUTCMonth() + 1)}`
      );
    }

    if (!rowIsInAvailableLogRange(curValues[trimToRowIdx])) {
      hasPreviousLogInRange = false;

    } else {
      do {
        trimToRowIdx++;
        if (trimToRowIdx >= curValues.length || !rowIsInAvailableLogRange(curValues[trimToRowIdx])) {
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
  let lastImportedIsWithinOneWeek = false;
  let monthsToImport = apiResponseMonthsAvailable;
  if (hasPreviousLogInRange) {
    const [lastImportedLogTimeStr, lastImportedLogNumStr, lastImportedLogActionStr]
      = getRowProperties(logHeaderRow, curValues[1], [logSheetInfo.header.datetime, logSheetInfo.header.total, logSheetInfo.header.reasonId]);
    lastImportedLogTime = Date.parse(ensureApiTime(lastImportedLogTimeStr));
    lastImportedLogNum = parseInt(lastImportedLogNumStr);
    lastImportedLogAction = parseInt(lastImportedLogActionStr);
    lastImportedIsWithinOneWeek = Date.now() - lastImportedLogTime < 1000 * 60 * 60 * 24 * 7;

    // remove months already imported
    const lastImportedLogTimeObj = new Date(lastImportedLogTime);
    const lastImportedLogMonthWithYear
      = `${lastImportedLogTimeObj.getFullYear()}-${padMonth(lastImportedLogTimeObj.getMonth() + 1)}`;
    monthsToImport = apiResponseMonthsAvailable.filter((_, i) => apiResponseMonthsAvailableWithYear[i] >= lastImportedLogMonthWithYear);
  }

  let newRows = [];
  let matchedLastLog = false;
  let cachedDueToTimeOut = false;

  const processEntries = (entries: LedgerLogEntry[], rows: string[][]) => addEntriesToRows(
    entries, logHeaderRow, rows, lastImportedLogTime,
    (entry: LedgerLogEntry) => {
      const { time, num, action_id } = entry;
      return Date.parse(time) === lastImportedLogTime
        && num === lastImportedLogNum
        && action_id === lastImportedLogAction;
    }
  );

  goingThroughAllMonths:
  for (let curMonthIdx = 0; curMonthIdx < monthsToImport.length; curMonthIdx++) {
    params.month = monthsToImport[curMonthIdx].toString();

    let curMonthRows = [];
    let startingPage = 1;

    // check if there is a cached data first
    const curMonthCacheSheetName = `${LOG_CACHE_PREFIX}:${logSheetInfo.sheetName}:${apiResponseMonthsAvailableWithYear[curMonthIdx]}`;
    let curMonthCacheSheet = SpreadsheetApp.getActive().getSheetByName(curMonthCacheSheetName);
    if (curMonthCacheSheet) {
      const cachedValues = curMonthCacheSheet.getDataRange().getValues();

      // check if the cached data is still valid
      const baselineRow = cachedValues[1];
      let curCheckingRow = 1;
      let curCheckingPage = 1;

      checkingCachedData:
      while (true) {
        serverDivideSpecificOp(serverDivide, {
          cn: () => { (params as LedgerParamsCn).page = (curCheckingPage).toString(); },
          os: () => { (params as LedgerParamsOs).current_page = (curCheckingPage).toString(); }
        });

        const fetchedRows = [];
        processEntries((requestApiResponse(endpoint, params, cookies) as LedgerApiResponse).data.list, fetchedRows);

        for (const fetchedRow of fetchedRows) {
          if (!isSameRowValue(cachedValues[curCheckingRow], fetchedRow, sheetTimeToStrProc)) {
            // failed, data has changed, start over
            break checkingCachedData;
          }
          if (!isSameRowValue(baselineRow, fetchedRow, sheetTimeToStrProc)) {
            // succeed, more than baselineRow matches, use the cached data
            startingPage = parseInt(cachedValues[0][0]);
            curMonthRows = cachedValues.slice(1);
            break checkingCachedData;
          }

          curCheckingRow++;
        }

        curCheckingPage++;
      }
    }

    let fetchedDataArray: LedgerLogData[]
      = [...Array(lastImportedIsWithinOneWeek ? 2 : LEDGER_FETCH_MULTI).fill(null)];
    let processedUpToIdx = 0, foundEnd = false;

    goingThroughCurMonth:
    while (true) {
      // cache fetched data if time runs out
      const curRunTime = Date.now() - importStartTime;
      if (curRunTime > LEDGER_RUN_TIME_LIMIT) {
        cachedDueToTimeOut = true;

        if (!curMonthCacheSheet) {
          curMonthCacheSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(curMonthCacheSheetName);
        }
        curMonthCacheSheet.getRange(1, 1).setValues([[startingPage + processedUpToIdx]]);
        curMonthCacheSheet.getRange(2, 1, curMonthRows.length, curMonthRows[0].length).setValues(curMonthRows);

        break goingThroughAllMonths;
      }

      // popluate requests with responses not yet fetched
      const requests: GoogleAppsScript.URL_Fetch.URLFetchRequest[] = [];
      for (let i = processedUpToIdx; i < fetchedDataArray.length; i++) {
        if (!fetchedDataArray[i]) {
          serverDivideSpecificOp(serverDivide, {
            cn: () => { (params as LedgerParamsCn).page = (i + startingPage).toString(); },
            os: () => { (params as LedgerParamsOs).current_page = (i + startingPage).toString(); }
          });
          requests.push(getApiRequest(endpoint, params, cookies));
        }
      }

      if (requests.length === 0) {
        // process remaining
        while (processedUpToIdx < fetchedDataArray.length) {
          const stoppingMatched = processEntries(fetchedDataArray[processedUpToIdx].list, curMonthRows);
          if (stoppingMatched) {
            matchedLastLog = true;
            break goingThroughCurMonth;
          }
          processedUpToIdx++;
        }
        break goingThroughCurMonth;  // not matchedLastLog
      }

      const curResponses = UrlFetchApp.fetchAll(requests);

      // parse and collect fetched pages
      let fetchSucceededCount = 0;
      for (const response of curResponses) {
        const parsed: LedgerApiResponse = JSON.parse(response.getContentText());

        if (parsed.retcode === 0) {
          const i = serverDivideSpecificOp(serverDivide, {
            cn: () => (parsed.data as LedgerLogDataCn).page - startingPage,
            os: () => (parsed.data as LedgerLogDataOs).current_page - startingPage,
          });
          if (parsed.data.list.length === 0) {
            fetchedDataArray = fetchedDataArray.slice(0, i);
            foundEnd = true;
            break;
          }

          fetchedDataArray[i] = parsed.data;
          fetchSucceededCount++;

        } else if (serverDivideSpecificOp(serverDivide, {
          cn: () => parsed.retcode !== LEDGER_ERROR_RESPONSE_TOO_MANY_ATTEMPTS_CN.retcode,
          os: () => parsed.retcode !== LEDGER_ERROR_RESPONSE_TOO_MANY_ATTEMPTS_OS.retcode,
        })) {
          throw new Error(`api request failed with retcode "${parsed.retcode}", msg: "${parsed.message}"`);
        }
      }

      // process fetched pages
      while (processedUpToIdx < fetchedDataArray.length && fetchedDataArray[processedUpToIdx]) {
        const stoppingMatched = processEntries(fetchedDataArray[processedUpToIdx].list, curMonthRows);
        if (stoppingMatched) {
          matchedLastLog = true;
          break goingThroughCurMonth;
        }
        processedUpToIdx++;
      }

      if (!foundEnd) {
        // extend fetchedDataArray to fill LEDGER_FETCH_MULTI limit
        fetchedDataArray.push(...Array(fetchSucceededCount).fill(null));
      }
    }

    newRows = curMonthRows.concat(newRows);
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

  if (cachedDueToTimeOut) {
    throw new Error(`RUN IMPORT AGAIN (MORE TIME NEEDED)`);
  }
}

const getPrimogemLog = () => writeImServiceLogToSheet(PRIMOGEM_SHEET_INFO);
const getCrystalLog = () => writeImServiceLogToSheet(CRYSTAL_SHEET_INFO);
const getResinLog = () => writeImServiceLogToSheet(RESIN_SHEET_INFO);
const getArtifactLog = () => writeImServiceLogToSheet(ARTIFACT_SHEET_INFO);
const getWeaponLog = () => writeImServiceLogToSheet(WEAPON_SHEET_INFO);
const getMoraLog = () => writeLedgerLogToSheet(MORA_SHEET_INFO);

const LOG_RANGES = {
  "Primogem Log": { "range_status": "E26", "range_toggle": "E19", "range_dashboard_length": "C15" },
  "Crystal Log": { "range_status": "E27", "range_toggle": "E20", "range_dashboard_length": "C20" },
  "Resin Log": { "range_status": "E28", "range_toggle": "E21", "range_dashboard_length": "C25" },
  "Artifact Log": { "range_status": "E29", "range_toggle": "E22", "range_dashboard_length": "C35" },
  "Weapon Log": { "range_status": "E46", "range_toggle": "E45", "range_dashboard_length": "C40" },
  "Mora Log": { "range_status": "E39", "range_toggle": "E35", "range_dashboard_length": "C30" }
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
        } else if (logName == SHEET_NAME_ARTIFACT_LOG) {
          getArtifactLog();
        } else if (logName == SHEET_NAME_WEAPON_LOG) {
          getWeaponLog();
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

const padMonth = (month: number) => month.toString().padStart(2, '0');
