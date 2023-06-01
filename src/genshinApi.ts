// for license and source, visit https://github.com/3096/primorina

type ServerDivide = "cn" | "os";
type RegionCode = "cn_gf01" | "cn_qd01" | "os_usa" | "os_euro" | "os_asia" | "os_cht";
type LocaleCode =
  "en-us" | "de-de" | "fr-fr" | "es-es" | "zh-tw" | "zh-cn" | "id-id" |
  "ja-jp" | "vi-vn" | "ko-kr" | "pt-pt" | "th-th" | "ru-ru";

const API_DOMAINS_BY_SERVER_DIVIDE = {
  cn: "hk4e-api.mihoyo.com",
  os: "hk4e-api-os.hoyoverse.com",
}

const KNOWN_DOMAIN_LIST = [
  { domain: "user.mihoyo.com", serverDivide: "cn" },
  { domain: "account.mihoyo.com", serverDivide: "os" },
  { domain: "account.hoyoverse.com", serverDivide: "os" },

  { domain: "webstatic.mihoyo.com", serverDivide: "cn" },
  { domain: "webstatic-sea.mihoyo.com", serverDivide: "os" },
  { domain: "webstatic-sea.hoyoverse.com", serverDivide: "os" },

  { domain: "hk4e-api.mihoyo.com", serverDivide: "cn" },
  { domain: "hk4e-api-os.mihoyo.com", serverDivide: "os" },
  { domain: "hk4e-api-os.hoyoverse.com", serverDivide: "os" },
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

interface Cookies {
  [name: string]: string;
}

interface Params {
  [name: string]: string;
}


// imService/ulog

interface ImServiceParams extends Params {
  authkey_ver: string,
  sign_type: string,
  auth_appid: string,
  device_type: string,
  authkey: string,
  lang: LocaleCode,
  size: string,
  end_id?: string,
}

function getImServiceDefaultQueryParams(authkey: string, locale: LocaleCode = "en-us"): ImServiceParams {
  return {
    authkey_ver: "1",
    sign_type: "2",
    auth_appid: "webview_gacha",
    device_type: "pc",
    authkey: authkey,
    lang: "en-us",
    size: "20",
  };
};

interface ImServiceLogEntry {
  id: string,
  uid: string,
  time: string,
  datetime: string,
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


// ledger

interface LedgerCookieCn extends Cookies {
  account_id: string, cookie_token: string
}

interface LedgerCookieOs extends Cookies {
  ltoken: string, ltuid: string
}

interface LedgerParams extends Params {
  type: string,
  month?: string,
}

interface LedgerParamsCn extends LedgerParams {
  limit: string,
  bind_region?: string,
  page?: string,
}

interface LedgerParamsOs extends LedgerParams {
  region: string,
  lang: LocaleCode,
  uid: string,
  current_page?: string,
}

const LEDGER_GET_DEFAULT_QUERY_PARAMS: {
  [serverDivide in ServerDivide]: (regionCode: RegionCode) => LedgerParams
} & {
  cn: (regionCode: RegionCode) => LedgerParamsCn,
  os: (regionCode: RegionCode) => LedgerParamsOs,
} = {
  cn: (regionCode: RegionCode) => {
    return { type: "2", limit: "100", bind_region: regionCode };
  },
  os: (regionCode: RegionCode, locale: LocaleCode = "en-us") => {
    return { type: "2", region: regionCode, lang: locale, uid: "1" };
  },
};

interface LedgerLogEntry {
  time: string,
  num: number,
  action_id: number,
  action: string,
}

interface LedgerLogData {
  optional_month: number[],
  data_month: number,
  region: string,
  uid: number,
  nickname: string,
  list: LedgerLogEntry[],
}

interface LedgerLogDataCn extends LedgerLogData {
  page: number,
}

interface LedgerLogDataOs extends LedgerLogData {
  current_page: number,
}

interface LedgerApiResponse {
  retcode: number,
  message: string,
  data: LedgerLogData,
}

const LEDGER_ERROR_RESPONSE_TOO_MANY_ATTEMPTS_CN: LedgerApiResponse = {
  "data": null, "message": "操作太频繁，请稍后再试", "retcode": -3
};

const LEDGER_ERROR_RESPONSE_TOO_MANY_ATTEMPTS_OS: LedgerApiResponse = {
  "data": null, "message": "Too many attempts. Try again later", "retcode": -500004
};

function getApiRequestUrlAndFetchParams(endpoint: string, params: Params, cookies: Cookies = null): {
  url: string, fetchParams: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions
} {
  const url = getUrlWithParams(endpoint, params);
  const fetchParams = {};
  if (cookies) {
    fetchParams["headers"] = {
      "Cookie": Object.entries(cookies)
        .filter(([, value]) => value !== null)
        .map(([name, value]) => `${name}=${value}`)
        .join("; ")
    };
  }

  return { url, fetchParams };
}

function requestApiResponse(endpoint: string, params: Params, cookies: Cookies = null) {
  const { url, fetchParams } = getApiRequestUrlAndFetchParams(endpoint, params, cookies);
  const response = JSON.parse(UrlFetchApp.fetch(url, fetchParams).getContentText());
  if (response.retcode !== 0) {
    throw new Error(`api request failed with retcode "${response.retcode}", msg: "${response.message}"`);
  }

  return response;
}

function getApiRequest(endpoint: string, params: Params, cookies: Cookies = null) {
  const { url, fetchParams } = getApiRequestUrlAndFetchParams(endpoint, params, cookies);
  const request = UrlFetchApp.getRequest(url, fetchParams);
  delete request.headers["X-Forwarded-For"];  // idk what this is but it messes things up
  return request;
}

function getReasonMap(config = getConfig()): Map<string, number> {
  const LANG_MAP_URL = `https://webstatic.hoyoverse.com/admin/mi18n/hk4e_global/m02251421001311/m02251421001311-${config.languageCode}.json`;
  const REASON_PREFIX = "selfinquiry_general_reason_";

  const langMap = JSON.parse(UrlFetchApp.fetch(LANG_MAP_URL).getContentText());

  const result = new Map<string, number>();
  for (const key in langMap) {
    if (!key.includes(REASON_PREFIX)) continue;

    const reasonId = parseInt(key.substring(REASON_PREFIX.length));
    result.set(langMap[key], reasonId);
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

function getUrlWithParams(urlEndpoint: string, params: Params) {
  return urlEndpoint + "?" + Object.entries(params).map(([key, value]) => `${key}=${value}`).join("&");
}

function getServerTimeAsUtcNow(regionCode: RegionCode) {
  return new Date(Date.now() + REGION_INFO[regionCode].timezone * 3600000);
}

// ApiTimeStr = "YYYY-MM-dd HH:mm:ss"

function sheetDateToApiTimeStr(sheetDate: Date) {  // only use this on date read from the sheet
  return Utilities.formatDate(
    sheetDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "YYYY-MM-dd HH:mm:ss");
}

function ensureApiTime(timeStr: string) {
  // regex check if the string is in the format of "YYYY-MM-dd HH:mm:ss"
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(timeStr)) {
    return timeStr;
  }
  return sheetDateToApiTimeStr(new Date(timeStr));
}

function getApiTimeAsServerTimeAsUtc(timeStr: string) {
  return new Date(ensureApiTime(timeStr).replace(" ", "T") + "Z");
}
