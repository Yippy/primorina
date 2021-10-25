// for license and source, visit https://github.com/3096/primorina

type ServerDivide = "cn" | "os";

const API_DOMAINS_BY_SERVER_DIVIDE = {
  cn: "hk4e-api.mihoyo.com",
  os: "hk4e-api-os.mihoyo.com",
}

const KNOWN_DOMAIN_LIST = [
  { domain: "user.mihoyo.com", serverDivide: "cn" },
  { domain: "account.mihoyo.com", serverDivide: "os" },

  { domain: "webstatic.mihoyo.com", serverDivide: "cn" },
  { domain: "webstatic-sea.mihoyo.com", serverDivide: "os" },

  { domain: "hk4e-api.mihoyo.com", serverDivide: "cn" },
  { domain: "hk4e-api-os.mihoyo.com", serverDivide: "os" },
];

const API_PARAM_AUTH_KEY = "authkey";
const API_PARAM_LANG = "lang";
const API_PARAM_REGION = "region";
const API_PARAM_SIZE = "size";
const API_END_ID = "end_id";
const API_PARAM_CURRENT_PAGE = "current_page";
const API_PARAM_MONTH = "month";

interface Cookies {
  [name: string]: string;
}

const getDefaultQueryParams = () => new Map([
  ["authkey_ver", "1"],
  ["sign_type", "2"],
  ["auth_appid", "webview_gacha"],
  ["device_type", "pc"],
]);

const getDefaultQueryParamsForHoYoLab = () => new Map([
  ["type", "2"],
  ["uid", "1"],
]);

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

interface LogEntryHoYo {
  time: string,
  num: string,
  action_id: string,
  action: string,
}

interface LogDataHoYo {
  optional_month: [],
  current_page: string,
  data_month: string,
  region: string,
  uid: string,
  nickname: string,
  list: LogEntryHoYo[],
}

interface ApiResponseHoYo {
  retcode: number,
  message: string,
  data: LogDataHoYo,
}

function requestApiResponse(endpoint: string, params: Map<string, string>, cookies: Cookies = null) {
  const url = getUrlWithParams(endpoint, params);
  const fetchParams = {};
  if (cookies) {
    fetchParams["headers"] =
      { "Cookie": Object.entries(cookies).map(keyValuePair => `${keyValuePair[0]}=${keyValuePair[1]}`).join("; ") };
  }

  Logger.log(fetchParams);

  const response = JSON.parse(UrlFetchApp.fetch(url, fetchParams).getContentText());
  if (response.retcode !== 0) {
    throw new Error(`api request failed with retcode "${response.retcode}", msg: "${response.message}"`);
  }

  return response;
}

function requestApiResponseHoYo(endpoint: string, params: Map<string, string>) {
  const cookie: Cookies = { ltoken: ltokenInput, ltuid: ltuidInput };
  return requestApiResponse(endpoint, params, cookie) as ApiResponseHoYo;
}

function getReasonMap(config = getConfig()) {
  const LANG_MAP_URL = `https://mi18n-os.mihoyo.com/webstatic/admin/mi18n/hk4e_global/m02251421001311/m02251421001311-${config.languageCode}.json`;
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

function getApiEndpoint(logSheetInfo: ILogSheetInfo, serverDivide: ServerDivide) {
  return "https://" + API_DOMAINS_BY_SERVER_DIVIDE[serverDivide] + logSheetInfo.apiPath;
}

function getServerDivideFromUrl(url: string) {
  for (const curItem of KNOWN_DOMAIN_LIST) {
    if (url.includes(curItem.domain)) {
      return curItem.serverDivide as ServerDivide;
    }
  }
  throw new Error(`no know domain detected in url "${url}"`);
}

function getParamValueFromUrlQueryString(url: string, param: string) {
  const anchor = url.indexOf("#");
  if (anchor >= 0) {
    url = url.substring(0, anchor);
  }

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
