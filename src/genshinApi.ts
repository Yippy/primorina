// for license and source, visit https://github.com/3096/primorina

type ServerDivide = "cn" | "os";
type RegionCode = "cn_gf01" | "cn_qd01" | "os_usa" | "os_euro" | "os_asia" | "os_cht";

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

const REGION_INFO: {
  [regionCode in RegionCode]: { serverDivide: ServerDivide, timezone: number }
} = {
  "cn_gf01": { serverDivide: "cn", timezone: +8, },  // 天空岛
  "cn_qd01": { serverDivide: "cn", timezone: +8, },  // 世界树
  "os_usa": { serverDivide: "os", timezone: -5, },  // America
  "os_euro": { serverDivide: "os", timezone: +1, },  // Europe
  "os_asia": { serverDivide: "os", timezone: +8, },  // Asia
  "os_cht": { serverDivide: "os", timezone: +8, },  // TW, HK, MO
}

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

const getImServiceDefaultQueryParams = () => new Map([
  ["authkey_ver", "1"],
  ["sign_type", "2"],
  ["auth_appid", "webview_gacha"],
  ["device_type", "pc"],
]);

const getLedgerDefaultQueryParams = () => new Map([
  ["type", "2"],
  ["uid", "1"],
]);

interface ImServiceLogEntry {
  id: string,
  uid: string,
  time: string,
  add_num: string,
  reason: string,
}

interface ImServiceLogData {
  end_id: string,
  size: string,
  region: string,
  uid: string,
  nickname: string,
  list: ImServiceLogEntry[],
}

interface ImServiceApiResponse {
  retcode: number,
  message: string,
  data: ImServiceLogData,
}

interface LedgerLogEntry {
  time: string,
  num: number,
  action_id: number,
  action: string,
}

interface LedgerLogData {
  optional_month: number[],
  current_page: string,
  data_month: string,
  region: string,
  uid: string,
  nickname: string,
  list: LedgerLogEntry[],
}

interface LedgerApiResponse {
  retcode: number,
  message: string,
  data: LedgerLogData,
}

function requestApiResponse(endpoint: string, params: Map<string, string>, cookies: Cookies = null) {
  const url = getUrlWithParams(endpoint, params);
  const fetchParams = {};
  if (cookies) {
    fetchParams["headers"] =
      { "Cookie": Object.entries(cookies).map(keyValuePair => `${keyValuePair[0]}=${keyValuePair[1]}`).join("; ") };
  }

  const response = JSON.parse(UrlFetchApp.fetch(url, fetchParams).getContentText());
  if (response.retcode !== 0) {
    throw new Error(`api request failed with retcode "${response.retcode}", msg: "${response.message}"`);
  }

  return response;
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
  return "https://" + API_DOMAINS_BY_SERVER_DIVIDE[serverDivide] + logSheetInfo.apiPaths[serverDivide];
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

function getServerTimeAsUtcNow(regionCode: RegionCode) {
  return new Date(Date.now() + REGION_INFO[regionCode].timezone * 3600000);
}

// "YYYY-MM-DD HH:MM:SS"
function getApiTimeAsServerTimeAsUtc(apiTimeStr: string) {
  return new Date(apiTimeStr.replace(" ", "T") + "Z");
}
