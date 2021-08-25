// for license and source, visit https://github.com/3096/primorina

const SHEET_NAME_CONFIG = "Configs";
const SHEET_NAME_PRIMOGEM_LOG = "Primogem Log";


interface Config {
  authKeyUrl: string,
  lang: string,
}

function getConfig() {
  const values = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_CONFIG).getDataRange().getValues();
  const config: Config = Object.fromEntries(values);
  return config;
}


type ServerDivide = "cn" | "os";

function getServerDivideFromUrl(url: string) {
  const KNOWN_DOMAIN_LIST = [
    { domain: "user.mihoyo.com", serverDivide: "cn" },
    { domain: "account.mihoyo.com", serverDivide: "os" },

    { domain: "webstatic.mihoyo.com", serverDivide: "cn" },
    { domain: "webstatic-sea.mihoyo.com", serverDivide: "os" },

    { domain: "hk4e-api.mihoyo.com", serverDivide: "cn" },
    { domain: "hk4e-api-os.mihoyo.com", serverDivide: "os" },
  ]

  for (const curItem of KNOWN_DOMAIN_LIST) {
    if (url.includes(curItem.domain)) {
      return curItem.serverDivide as ServerDivide;
    }
  }
}

const API_DOMAINS_BY_SERVER_DIVIDE = {
  cn: "hk4e-api.mihoyo.com",
  os: "hk4e-api-os.mihoyo.com",
}

function getApiEndpoint(logSheetInfo: ILogSheetInfo, serverDivide: ServerDivide) {
  return "https://" + API_DOMAINS_BY_SERVER_DIVIDE[serverDivide] + logSheetInfo.apiPath;
}


const API_PARAM_AUTH_KEY = "authkey";
const API_PARAM_LANG = "lang";
const API_PARAM_SIZE = "size";
const API_END_ID = "end_id";

function getDefaultQueryParams() {
  return new Map<string, string>([
    ["authkey_ver", "1"],
    ["sign_type", "2"],
    ["auth_appid", "webview_gacha"],
    ["device_type", "pc"],
  ]);
}

function getParamValueFromUrlQuesryStryng(url: string, param: string) {
  const start = url.indexOf(param + "=") + param.length + 1;
  const end = url.indexOf("&", start);

  if (start < 0) {
    throw new Error(`cannot find param "${param}" in "${url}"`);
  }
  if (end < 0) {
    return url.substring(start);
  }

  return url.substring(start, end);
}

function getUrlWithParams(urlEndpoint: string, params: Map<string, string>) {
  let result = urlEndpoint + "?";
  for (const entry of params.entries()) {
    if (entry[1] && entry[0]) {
      result += entry[0] + "=" + entry[1] + "&";
    }
  }
  return result.slice(0, -1);
}


interface ILogSheetInfo {
  sheetName: string,
  apiPath: string
}

const PRIMOGEM_SHEET_INFO: ILogSheetInfo = {
  sheetName: SHEET_NAME_PRIMOGEM_LOG,
  apiPath: "/ysulog/api/getPrimogemLog",
}

// const CRYSTAL_SHEET_INFO: ILogSheetInfo = {
// sheetName: ,
// apiPath: "/ysulog/api/getPrimogemLog",
// }

interface LogEntry {
  id: string,
  uid: string,
  time: string,
  add_num: string,
  reason: string,
}

interface LogData {
  end_id: string,
  size: string,
  region: string,
  uid: string,
  nickname: string,
  list: LogEntry[],
}

interface ApiResponse {
  retcode: number,
  message: string,
  data: LogData,
}

function requestApiResponse(endpoint: string, params: Map<string, string>) {
  const response: ApiResponse = JSON.parse(UrlFetchApp.fetch(getUrlWithParams(endpoint, params)).getContentText());
  if (response.retcode !== 0) {
    throw new Error(`api request failed with retcode "${response.retcode}", msg: "${response.message}"`);
  }
  return response;
}


function getReasonMap() {
  const config = getConfig();
  const LANG_MAP_URL = `https://mi18n-os.mihoyo.com/webstatic/admin/mi18n/hk4e_global/m02251421001311/m02251421001311-${config.lang}.json`;
  const REASON_PREFIX = "selfinquiry_general_reason_";

  const langMap = JSON.parse(UrlFetchApp.fetch(LANG_MAP_URL).getContentText());

  const result = new Map<number, string>();
  for (const key in langMap) {
    if (!key.includes(REASON_PREFIX)) continue;

    const reasonId = parseInt(key.substring(REASON_PREFIX.length));
    result.set(reasonId, langMap[key]);
  }
  return result;
}
const REASON_MAP = getReasonMap();


function writeLogToSheet(logSheetInfo: ILogSheetInfo) {
  const config = getConfig();
  const endpoint = getApiEndpoint(logSheetInfo, getServerDivideFromUrl(config.authKeyUrl));

  const params = getDefaultQueryParams();
  params.set(API_PARAM_AUTH_KEY, getParamValueFromUrlQuesryStryng(config.authKeyUrl, API_PARAM_AUTH_KEY));
  params.set(API_PARAM_LANG, config.lang);
  params.set(API_PARAM_SIZE, "20");

  const logSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_PRIMOGEM_LOG);
  const curValues = logSheet.getDataRange().getValues();
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
          case "detail":
            const reasonId = parseInt(log.reason);
            newRow.push(REASON_MAP.get(reasonId));
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

  newRows.push(...curValues.slice(1));
  logSheet.getRange(2, 1, newRows.length, LOG_HEADER_ROW.length).setValues(newRows);
}

const getPrimogemLog = () => writeLogToSheet(PRIMOGEM_SHEET_INFO);
