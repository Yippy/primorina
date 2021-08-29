/********** Yippym **************/
const SHEET_NAME_PRIMOGEM_LOG = "Primogem Log";
const SHEET_NAME_CRYSTAL_LOG = "Crystal Log";
var SCRIPT_VERSION = "v0.0.2";
var sheetSourceId = '1p-SkTsyzoxuKHqqvCJSUCaFBUmxd5uEEvCtb7bAqfDk';
var nameOfLogHistorys = [SHEET_NAME_PRIMOGEM_LOG, SHEET_NAME_CRYSTAL_LOG];

function displayAbout() {
  var titleString = "About";
  var htmlString = "Script version: " + SCRIPT_VERSION;
  var widthSize = 500;
  var heightSize = 400;

  var htmlOutput = HtmlService
    .createHtmlOutput(htmlString)
    .setWidth(widthSize) //optional
    .setHeight(heightSize); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, titleString);
}

function moveToDashboardSheet() {
  moveToSheetByName("Dashboard");
}
function moveToSettingsSheet() {
  moveToSheetByName("Settings");
}
function moveToChangelogSheet() {
  moveToSheetByName("Changelog");
}
function moveToReadmeSheet() {
  moveToSheetByName("README");
}
function moveToPrimogemLogSheet() {
  moveToSheetByName("Primogem Log");
}
function moveToCrystalLogSheet() {
  moveToSheetByName("Crystal Log");
}

function moveToSheetByName(nameOfSheet) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(nameOfSheet);
  if (sheet) {
    sheet.activate();
  } else {
    const title = "Error";
    const message = "Unable to find sheet named '" + nameOfSheet + "'.";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function displayReadme() {
  var sheetSource = SpreadsheetApp.openById(sheetSourceId);
  if (sheetSource) {
    // Avoid Exception: You can't remove all the sheets in a document.Details
    var placeHolderSheet = null;
    if (SpreadsheetApp.getActive().getSheets().length == 1) {
      placeHolderSheet = SpreadsheetApp.getActive().insertSheet();
    }
    var sheetToRemove = SpreadsheetApp.getActive().getSheetByName('README');
    if (sheetToRemove) {
      // If exist remove from spreadsheet
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToRemove);
    }
    var sheetREADMESource;

    // Add Language
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
    if (settingsSheet) {
      var languageFound = settingsSheet.getRange(2, 2).getValue();
      sheetREADMESource = sheetSource.getSheetByName("README" + "-" + languageFound);
    }
    if (sheetREADMESource) {
      // Found language
    } else {
      // Default
      sheetREADMESource = sheetSource.getSheetByName("README");
    }

    sheetREADMESource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('README');

    // Remove placeholder if available
    if (placeHolderSheet) {
      // If exist remove from spreadsheet
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(placeHolderSheet);
    }
    var sheetREADME = SpreadsheetApp.getActive().getSheetByName('README');
    // Refresh Contents Links
    var contentsAvailable = sheetREADME.getRange(13, 1).getValue();
    var contentsStartIndex = 15;

    for (var i = 0; i < contentsAvailable; i++) {
      var valueRange = sheetREADME.getRange(contentsStartIndex + i, 1).getValue();
      var formulaRange = sheetREADME.getRange(contentsStartIndex + i, 1).getFormula();
      // Display for user, doesn't do anything
      sheetREADME.getRange(contentsStartIndex + i, 1).setFormula(formulaRange);

      // Grab URL RichTextValue from Source
      const range = sheetREADMESource.getRange(contentsStartIndex + i, 1);
      const RichTextValue = range.getRichTextValue().getRuns();
      const res = RichTextValue.reduce((ar, e) => {
        const url = e.getLinkUrl();
        if (url) ar.push(url);
        return ar;
      }, []);
      //  Convert to string
      var resString = res + "";
      var arrayString = resString.split("=");
      if (arrayString.length > 1) {
        var text = arrayString[2];
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(valueRange)
          .setLinkUrl("#gid=" + sheetREADME.getSheetId() + 'range=' + text)
          .build();
        sheetREADME.getRange(contentsStartIndex + i, 1).setRichTextValue(richText);
      }
    }
    reorderSheets();
    SpreadsheetApp.getActive().setActiveSheet(sheetREADME);
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function reorderSheets() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
  if (settingsSheet) {
    var sheetsToSort = settingsSheet.getRange(11, 2, 11, 1).getValues();

    for (var i = 0; i < sheetsToSort.length; i++) {
      var sheetName = sheetsToSort[i][0];
      if (sheetName != "") {
        var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (sheet) {
          SpreadsheetApp.getActive().setActiveSheet(sheet);
          var position = i + 1;
          if (position >= SpreadsheetApp.getActive().getNumSheets()) {
            position = SpreadsheetApp.getActive().getNumSheets();
          }
          SpreadsheetApp.getActive().moveActiveSheet(position);
        }
      }
    }
  }
}

var dashboardEditRange = [
  "I5", // Status cell
  "AB28", // Document Version
  "AT40", // Current Document Version
  "T29", // Document Status
  "AV1", // Name of drop down 1 (import)
  "AV2", // Name of drop down 2 (auto import)
  "AG14", // Selection
  "AG18" // URL
];

// Cells that needs refreshing
var dashboardRefreshRange = [
];

function quickUpdate() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  if (dashboardSheet) {
    dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Running script, please wait.");
    dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("yellow").setFontWeight("bold");
  }
  if (dashboardSheet) {
    if (settingsSheet) {
      var isLoading = settingsSheet.getRange(9, 7).getValue();

      if (isLoading) {
        var counter = settingsSheet.getRange(9, 8).getValue();
        if (counter > 0) {
          counter++;
          settingsSheet.getRange(9, 8).setValue(counter);
        } else {
          settingsSheet.getRange(9, 8).setValue(1);
        }
        if (counter > 2) {
          // Bypass message - for people with broken update wanting force update
        } else {
          var message = 'Still updating';
          var title = 'Quick Update already started, the number of time you requested is ' + counter + '. If you want to force an quick update due to an error happened during update, proceed in calling "Update Item" one more try.';
          SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
          return;
        }
      } else {
        settingsSheet.getRange(9, 7).setValue(true);
        settingsSheet.getRange(9, 8).setValue(1);
        settingsSheet.getRange("G10").setValue(new Date());
      }

      var changelogSheet = SpreadsheetApp.getActive().getSheetByName('Changelog');
      if (changelogSheet) {
        try {
          var sheetSource = SpreadsheetApp.openById(sheetSourceId);
          if (sheetSource) {
            // check latest logs to see anything new
            if (dashboardSheet) {
              var sheetAvailableSource = sheetSource.getSheetByName("Available");
              if (dashboardSheet) {
                var sourceDocumentVersion = sheetAvailableSource.getRange("E1").getValues();
                var currentDocumentVersion = dashboardSheet.getRange(dashboardEditRange[2]).getValues();
                dashboardSheet.getRange(dashboardEditRange[1]).setValue(sourceDocumentVersion);
                if (sourceDocumentVersion > currentDocumentVersion) {
                  dashboardSheet.getRange(dashboardEditRange[3]).setValue("New Document Available, make a new copy");
                  dashboardSheet.getRange(dashboardEditRange[3]).setFontColor("red").setFontWeight("bold");
                } else {
                  dashboardSheet.getRange(dashboardEditRange[3]).setValue("Document is up-to-date");
                  dashboardSheet.getRange(dashboardEditRange[3]).setFontColor("green").setFontWeight("bold");
                }
              }
              var changesCheckRange = changelogSheet.getRange(2, 1).getValue();
              changesCheckRange = changesCheckRange.split(",");
              var lastDateChangeText;
              var lastDateChangeSourceText;
              var isChangelogTheSame = true;

              var sheetChangelogSource = sheetSource.getSheetByName("Changelog");
              for (var i = 0; i < changesCheckRange.length; i++) {
                var checkChangelogSource = sheetChangelogSource.getRange(changesCheckRange[i]).getValue();
                if (checkChangelogSource instanceof Date) {
                  lastDateChangeSourceText = Utilities.formatDate(checkChangelogSource, 'Etc/GMT', 'dd-MM-yyyy');
                }
                var checkChangelog = changelogSheet.getRange(changesCheckRange[i]).getValue();
                if (checkChangelog instanceof Date) {
                  lastDateChangeText = Utilities.formatDate(checkChangelog, 'Etc/GMT', 'dd-MM-yyyy');
                  if (lastDateChangeSourceText != lastDateChangeText) {
                    isChangelogTheSame = false;
                    break;
                  }
                } else {
                  if (checkChangelogSource != checkChangelog) {
                    isChangelogTheSame = false;
                    break;
                  }
                }
              }
              if (isChangelogTheSame) {
                dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: There is no changes from source");
                dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("green").setFontWeight("bold");
              } else {
                if (lastDateChangeText == lastDateChangeSourceText) {
                  dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Current Changelog has the same date " + lastDateChangeText + " but isn't the same notes to source. Please run 'Update Items'.");
                } else {
                  dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Current Changelog is " + lastDateChangeText + ", source is at " + lastDateChangeSourceText + ". Please run 'Update Items'.");
                }
                dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
              }
            }
          } else {
            if (dashboardSheet) {
              dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Unable to connect to source, try again next time");
              dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
            }
          }
        } catch (e) {
          if (dashboardSheet) {
            dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Unable to connect to source, try again next time.");
            dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
          }
        }
      } else {
        if (dashboardSheet) {
          dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Missing 'Changelog' sheet in this Document, unable to compare to source. Please run 'Update Items'.");
          dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
        }
      }
    } else {
      if (dashboardSheet) {
        dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Missing 'Settings' sheet in this Document, make a new copy as this Document has important sheet missing.");
        dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
      }
    }
  }
  var currentSheet = SpreadsheetApp.getActive().getActiveSheet();
  reorderSheets();
  SpreadsheetApp.getActive().setActiveSheet(currentSheet);
  // Update Settings
  settingsSheet.getRange(9, 7).setValue(false);
  settingsSheet.getRange("H10").setValue(new Date());
}

function importButtonScript() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  if (dashboardSheet && settingsSheet) {
    var userImportSelection = dashboardSheet.getRange(dashboardEditRange[4]).getValue();
    var importSelectionText = dashboardSheet.getRange(dashboardEditRange[6]).getValue();
    var urlInput = dashboardSheet.getRange(dashboardEditRange[7]).getValue();
    dashboardSheet.getRange(dashboardEditRange[7]).setValue(""); //Clear input
    if (userImportSelection == importSelectionText) {
      settingsSheet.getRange("D6").setValue(urlInput);
      importDataManagement();
    } else {
      settingsSheet.getRange("D17").setValue(urlInput);
      importFromAPI();
    }
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Unable to find 'Dashboard' or 'Settings'", "Missing Sheets");
  }
}

function importDataManagement() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  var userImportInput = settingsSheet.getRange("D6").getValue();
  var userImportStatus = settingsSheet.getRange("E7").getValue();
  var completeStatus = "COMPLETE";
  var wishHistoryNotDoneStatus = "NOT DONE";
  var wishHistoryDoneStatus = "DONE";
  var wishHistoryMissingStatus = "NOT FOUND";
  var message = "";
  var title = "";
  var statusMessage = "";
  var rowOfStatusWishHistory = 9;
  if (userImportStatus == completeStatus) {
    title = "Error";
    message = "Already done, you only need to run once";
  } else {
    if (userImportInput) {
      // Attempt to load as URL
      var importSource;
      try {
        importSource = SpreadsheetApp.openByUrl(userImportInput);
      } catch (e) {
      }
      if (importSource) {
      } else {
        // Attempt to load as ID instead
        try {
          importSource = SpreadsheetApp.openById(userImportInput);
        } catch (e) {
        }
      }
      if (importSource) {
        // Go through the available sheet list
        for (var i = 0; i < nameOfLogHistorys.length; i++) {
          var bannerImportSheet = importSource.getSheetByName(nameOfLogHistorys[i]);

          var numberOfRows = bannerImportSheet.getMaxRows() - 1;
          var range = bannerImportSheet.getRange(2, 1, numberOfRows, 5);

          if (bannerImportSheet && numberOfRows > 0) {
            var bannerSheet = SpreadsheetApp.getActive().getSheetByName(nameOfLogHistorys[i]);

            if (bannerSheet) {
              bannerSheet.getRange(2, 1, numberOfRows, 5).setValues(range.getValues());
              settingsSheet.getRange(rowOfStatusWishHistory + i, 5).setValue(wishHistoryDoneStatus);
            } else {
              settingsSheet.getRange(rowOfStatusWishHistory + i, 5).setValue(wishHistoryMissingStatus);
            }
          } else {
            settingsSheet.getRange(rowOfStatusWishHistory + i, 5).setValue(wishHistoryMissingStatus);
          }
        }
        var sourceSettingsSheet = importSource.getSheetByName("Settings");
        if (sourceSettingsSheet) {
          var language = sourceSettingsSheet.getRange("B2").getValue();
          if (language) {
            settingsSheet.getRange("B2").setValue(language);
          }
          var server = sourceSettingsSheet.getRange("B3").getValue();
          if (server) {
            settingsSheet.getRange("B3").setValue(server);
          }
          var calendarStartWeek = sourceSettingsSheet.getRange("B5").getValue();
          if (calendarStartWeek) {
            settingsSheet.getRange("B5").setValue(calendarStartWeek);
          }
        }

        title = "Complete";
        message = "Imported all rows in column Paste Value and Override";
        statusMessage = completeStatus;
      } else {
        title = "Error";
        message = "Import From URL or Spreadsheet ID is invalid";
        statusMessage = "Failed";
      }
    } else {
      title = "Error";
      message = "Import From URL or Spreadsheet ID is empty";
      statusMessage = "Failed";
    }

    settingsSheet.getRange("E7").setValue(statusMessage);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}
/**
* Update Item List
*/
function updateItemsList() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  var updateItemHasFailed = false;
  if (dashboardSheet) {
    dashboardSheet.getRange(dashboardEditRange[0]).setValue("Update Items: Running script, please wait.");
    dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("yellow").setFontWeight("bold");
  }
  var sheetSource = SpreadsheetApp.openById(sheetSourceId);
  // Check source is available
  if (sheetSource) {
    try {
      // attempt to load sheet from source, to prevent removing sheets first.
      var sheetAvailableSource = sheetSource.getSheetByName("Available");
      // Avoid Exception: You can't remove all the sheets in a document.Details
      var placeHolderSheet = null;
      if (SpreadsheetApp.getActive().getSheets().length == 1) {
        placeHolderSheet = SpreadsheetApp.getActive().insertSheet();
      }
      var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
      if (settingsSheet) {
        var isLoading = settingsSheet.getRange(5, 7).getValue();
        if (isLoading) {
          var counter = settingsSheet.getRange(5, 8).getValue();
          if (counter > 0) {
            counter++;
            settingsSheet.getRange(5, 8).setValue(counter);
          } else {
            settingsSheet.getRange(5, 8).setValue(1);
          }
          if (counter > 2) {
            // Bypass message - for people with broken update wanting force update
          } else {
            var message = 'Still updating';
            var title = 'Update already started, the number of time you requested is ' + counter + '. If you want to force an update due to an error happened during update, proceed in calling "Update Item" one more try.';
            SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
            return;
          }
        } else {
          settingsSheet.getRange(5, 7).setValue(true);
          settingsSheet.getRange(5, 8).setValue(1);
          settingsSheet.getRange("G6").setValue(new Date());
        }
      }
      // Remove sheets
      var listOfSheetsToRemove = [];
      var availableRangesValues = sheetAvailableSource.getRange(2, 1, sheetAvailableSource.getMaxRows() - 1, 1).getValues();
      var availableRanges = String(availableRangesValues).split(",");

      if (dashboardSheet) {
        var sourceDocumentVersion = sheetAvailableSource.getRange("E1").getValues();
        var currentDocumentVersion = dashboardSheet.getRange(dashboardEditRange[2]).getValues();
        dashboardSheet.getRange(dashboardEditRange[1]).setValue(sourceDocumentVersion);
        if (sourceDocumentVersion > currentDocumentVersion) {
          dashboardSheet.getRange(dashboardEditRange[3]).setValue("New Document Available, make a new copy");
          dashboardSheet.getRange(dashboardEditRange[3]).setFontColor("red").setFontWeight("bold");
        } else {
          dashboardSheet.getRange(dashboardEditRange[3]).setValue("Document is up-to-date");
          dashboardSheet.getRange(dashboardEditRange[3]).setFontColor("green").setFontWeight("bold");
        }
      }
      // Go through the available sheet list
      for (var i = 0; i < availableRanges.length; i++) {
        listOfSheetsToRemove.push(availableRanges[i]);
      }
      var listOfSheetsToRemoveLength = listOfSheetsToRemove.length;
      for (var i = 0; i < listOfSheetsToRemoveLength; i++) {
        var sheetNameToRemove = listOfSheetsToRemove[i];
        var sheetToRemove = SpreadsheetApp.getActive().getSheetByName(sheetNameToRemove);
        if (sheetToRemove) {

          // If exist remove from spreadsheet
          SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToRemove);
        }
      }

      // Put available sheet into current
      var skipRangeValues = sheetAvailableSource.getRange(2, 2, sheetAvailableSource.getMaxRows() - 1, 1).getValues();
      var skipRanges = String(skipRangeValues).split(",");
      var hiddenRangeValues = sheetAvailableSource.getRange(2, 3, sheetAvailableSource.getMaxRows() - 1, 1).getValues();
      var hiddenRanges = String(hiddenRangeValues).split(",");
      var settingsOptionRangeValues = sheetAvailableSource.getRange(2, 4, sheetAvailableSource.getMaxRows() - 1, 1).getValues();
      var settingsOptionRanges = String(settingsOptionRangeValues).split(",");

      for (var i = 0; i < availableRanges.length; i++) {
        var nameOfBanner = availableRanges[i];
        var isSkipString = skipRanges[i];
        var isHiddenString = hiddenRanges[i];
        var settingOptionString = settingsOptionRanges[i];
        var settingOptionNum = parseInt(settingOptionString);

        var sheetAvailableSelectionSource = sheetSource.getSheetByName(nameOfBanner);
        var storedSheet;
        if (isSkipString == "YES") {
          // skip - disabled by source
        } else {
          if (sheetAvailableSelectionSource) {
            if (settingOptionString == "" || settingOptionNum == 0) {
              //Enable without settings
              storedSheet = sheetAvailableSelectionSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(nameOfBanner);
            } else {
              // Check current setting has row
              if (settingOptionNum <= settingsSheet.getMaxRows()) {
                var checkEnabledRanges = settingsSheet.getRange(settingOptionNum, 2).getValue();
                if (checkEnabledRanges == "YES") {
                  storedSheet = sheetAvailableSelectionSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(nameOfBanner);
                } else {
                  storedSheet = null;
                }
              } else {
                // Sheet does not have this settings available
                storedSheet = null;
              }
            }
            if (storedSheet) {
              if (isHiddenString == "YES") {
                storedSheet.hideSheet();
              } else {
                storedSheet.showSheet();
              }
            }
          }
        }
      }

      // Remove placeholder if available
      if (placeHolderSheet) {
        // If exist remove from spreadsheet
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(placeHolderSheet);
      }

      reorderSheets();

      SpreadsheetApp.getActive().setActiveSheet(dashboardSheet);
      // Update Settings
      settingsSheet.getRange(5, 7).setValue(false);
      settingsSheet.getRange("H6").setValue(new Date());

    } catch (e) {
      var message = 'Unable to connect to source';
      var title = 'Error';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      updateItemHasFailed = true;
      settingsSheet.getRange(5, 7).setValue(false);
      settingsSheet.getRange("H6").setValue(new Date());
    }
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    updateItemHasFailed = true;
    settingsSheet.getRange(5, 7).setValue(false);
    settingsSheet.getRange("H6").setValue(new Date());
  }

  if (dashboardSheet) {
    if (updateItemHasFailed) {
      dashboardSheet.getRange(dashboardEditRange[0]).setValue("Update Items: Update Items has failed, please try again.");
      dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
    } else {
      dashboardSheet.getRange(dashboardEditRange[0]).setValue("Update Items: Successfully updated the Item list.");
      dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("green").setFontWeight("bold");
    }
  }
}
/********************************/
// Compiled using primorina 0.0.1 (TypeScript 4.3.5)
// for license and source, visit https://github.com/3096/primorina

var __values = (this && this.__values) || function (o) {
  var s = typeof Symbol === "function" && Symbol.iterator, m = s && o[s], i = 0;
  if (m) {
    return m.call(o);
  }
  if (o && typeof o.length === "number") return {
    next: function () {
      if (o && i >= o.length) o = void 0;
      return { value: o && o[i++], done: !o };
    }
  };
  throw new TypeError(s ? "Object is not iterable." : "Symbol.iterator is not defined.");
};

var __read = (this && this.__read) || function (o, n) {
  var m = typeof Symbol === "function" && o[Symbol.iterator];
  if (!m) {
    return o;
  }
  var i = m.call(o), r, ar = [], e;
  try {
    while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);
  } catch (error) {
    e = { error: error };
  } finally {
    try {
      if (r && !r.done && (m = i["return"])) m.call(i);
    } finally {
      if (e) {
        throw e.error;
      }
    }
  }
  return ar;
};

var __spreadArray = (this && this.__spreadArray) || function (to, from) {
  for (var i = 0, il = from.length, j = to.length; i < il; i++, j++) {
    to[j] = from[i];
  }
  return to;
};

var SHEET_NAME_CONFIG = "Settings";

function getServerDivideFromUrl(url) {
  var e_1, _a;
  var KNOWN_DOMAIN_LIST = [
    { domain: "user.mihoyo.com", serverDivide: "cn" },
    { domain: "account.mihoyo.com", serverDivide: "os" },
    { domain: "webstatic.mihoyo.com", serverDivide: "cn" },
    { domain: "webstatic-sea.mihoyo.com", serverDivide: "os" },
    { domain: "hk4e-api.mihoyo.com", serverDivide: "cn" },
    { domain: "hk4e-api-os.mihoyo.com", serverDivide: "os" },
  ];
  try {
    for (var KNOWN_DOMAIN_LIST_1 = __values(KNOWN_DOMAIN_LIST), KNOWN_DOMAIN_LIST_1_1 = KNOWN_DOMAIN_LIST_1.next(); !KNOWN_DOMAIN_LIST_1_1.done; KNOWN_DOMAIN_LIST_1_1 = KNOWN_DOMAIN_LIST_1.next()) {
      var curItem = KNOWN_DOMAIN_LIST_1_1.value;
      if (url.includes(curItem.domain)) {
        return curItem.serverDivide;
      }
    }
  } catch (e_1_1) {
    e_1 = { error: e_1_1 };
  } finally {
    try {
      if (KNOWN_DOMAIN_LIST_1_1 && !KNOWN_DOMAIN_LIST_1_1.done && (_a = KNOWN_DOMAIN_LIST_1["return"])) _a.call(KNOWN_DOMAIN_LIST_1);
    } finally {
      if (e_1) {
        throw e_1.error;
      }
    }
  }
}

var API_DOMAINS_BY_SERVER_DIVIDE = {
  cn: "hk4e-api.mihoyo.com",
  os: "hk4e-api-os.mihoyo.com"
};

function getApiEndpoint(logSheetInfo, serverDivide) {
  return "https://" + API_DOMAINS_BY_SERVER_DIVIDE[serverDivide] + logSheetInfo.apiPath;
}

var API_PARAM_AUTH_KEY = "authkey";
var API_PARAM_LANG = "lang";
var API_PARAM_SIZE = "size";
var API_END_ID = "end_id";

function getDefaultQueryParams() {
  return new Map([
    ["authkey_ver", "1"],
    ["sign_type", "2"],
    ["auth_appid", "webview_gacha"],
    ["device_type", "pc"],
  ]);
}

function getParamValueFromUrlQueryString(url, param) {
  const anchor = url.indexOf("#");
  if (anchor >= 0) {
    url = url.substring(0, anchor);
  }
  var start = url.indexOf(param + "=") + param.length + 1;
  var end = url.indexOf("&", start);
  if (start < 0) {
    throw new Error("cannot find param \"" + param + "\" in \"" + url + "\"");
  }
  if (end < 0) {
    return url.substring(start);
  }
  return url.substring(start, end);
}

function getUrlWithParams(urlEndpoint, params) {
  var e_2, _a;
  var result = urlEndpoint + "?";
  try {
    for (var _b = __values(params.entries()), _c = _b.next(); !_c.done; _c = _b.next()) {
      var entry = _c.value;
      if (entry[1] && entry[0]) {
        result += entry[0] + "=" + entry[1] + "&";
      }
    }
  } catch (e_2_1) {
    e_2 = { error: e_2_1 };
  } finally {
    try {
      if (_c && !_c.done && (_a = _b["return"])) _a.call(_b);
    } finally {
      if (e_2) {
        throw e_2.error;
      }
    }
  }
  return result.slice(0, -1);
}

function requestApiResponse(endpoint, params) {
  var response = JSON.parse(UrlFetchApp.fetch(getUrlWithParams(endpoint, params)).getContentText());
  if (response.retcode !== 0) {
    throw new Error("api request failed with retcode \"" + response.retcode + "\", msg: \"" + response.message + "\"");
  }
  return response;
}

var languageSettingsForImport = {
  "English": { "code": "en", "full_code": "en-us", "4_star": " (4-Star)", "5_star": " (5-Star)" },
  "German": { "code": "de", "full_code": "de-de", "4_star": " (4 Sterne)", "5_star": " (5 Sterne)" },
  "French": { "code": "fr", "full_code": "fr-fr", "4_star": " (4★)", "5_star": " (5★)" },
  "Spanish": { "code": "es", "full_code": "es-es", "4_star": " (4★)", "5_star": " (5★)" },
  "Chinese Traditional": { "code": "zh-tw", "full_code": "zh-tw", "4_star": " (四星)", "5_star": " (五星)" },
  "Chinese Simplified": { "code": "zh-cn", "full_code": "zh-cn", "4_star": " (四星)", "5_star": " (五星)" },
  "Indonesian": { "code": "id", "full_code": "id-id", "4_star": " (4★)", "5_star": " (5★)" },
  "Japanese": { "code": "ja", "full_code": "ja-jp", "4_star": " (★4)", "5_star": " (★5)" },
  "Vietnamese": { "code": "vi", "full_code": "vi-vn", "4_star": " (4 sao)", "5_star": " (5 sao)" },
  "Korean": { "code": "ko", "full_code": "ko-kr", "4_star": " (★4)", "5_star": " (★5)" },
  "Portuguese": { "code": "pt", "full_code": "pt-pt", "4_star": " (4★)", "5_star": " (5★)" },
  "Thai": { "code": "th", "full_code": "th-th", "4_star": " (4 ดาว)", "5_star": " (5 ดาว)" },
  "Russian": { "code": "ru", "full_code": "ru-ru", "4_star": " (4★)", "5_star": " (5★)" }
};

function getReasonMap() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  var selectionLanguage = settingsSheet.getRange("B2").getValue();
  var LANG_MAP_URL = "https://mi18n-os.mihoyo.com/webstatic/admin/mi18n/hk4e_global/m02251421001311/m02251421001311-" + languageSettingsForImport[selectionLanguage]["full_code"] + ".json";
  var REASON_PREFIX = "selfinquiry_general_reason_";
  var langMap = JSON.parse(UrlFetchApp.fetch(LANG_MAP_URL).getContentText());
  var result = new Map();
  for (var key in langMap) {
    if (!key.includes(REASON_PREFIX))
      continue;
    var reasonId = parseInt(key.substring(REASON_PREFIX.length));
    result.set(reasonId, langMap[key]);
  }
  return result;
}

var REASON_MAP = getReasonMap();

function writeLogToSheet(logSheetInfo) {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  var authKeyUrl = settingsSheet.getRange("D17").getValue();
  var selectionLanguage = settingsSheet.getRange("B2").getValue();
  var languageCode = languageSettingsForImport[selectionLanguage]["full_code"];
  var getURL = getServerDivideFromUrl(authKeyUrl);
  if (getURL == undefined || getURL == "") {
    SpreadsheetApp.getActiveSpreadsheet().toast("Invalid Auth Key URL", "Error");
    settingsSheet.getRange(bannerSettingsForImport[logSheetInfo.sheetName]['range_status']).setValue("No URL");
  } else {
    settingsSheet.getRange(bannerSettingsForImport[logSheetInfo.sheetName]['range_status']).setValue("Starting");
    var endpoint = getApiEndpoint(logSheetInfo, getURL);
    var params = getDefaultQueryParams();
    params.set(API_PARAM_AUTH_KEY, getParamValueFromUrlQueryString(authKeyUrl, API_PARAM_AUTH_KEY));
    params.set(API_PARAM_LANG, languageCode);
    params.set(API_PARAM_SIZE, "20");
    var logSheet = SpreadsheetApp.getActive().getSheetByName(logSheetInfo.sheetName);
    var curValues = logSheet.getDataRange().getValues();
    var previousRowLength = curValues.length;
    var LOG_HEADER_ROW = curValues[0];
    var ID_INDEX = LOG_HEADER_ROW.indexOf("id");
    var newRows = [];
    var stopAtId = curValues.length > 1 ? curValues[1][ID_INDEX] : null;
    var addLogsToNewRows = function (newLogs) {
      var e_3, _a, e_4, _b;
      try {
        for (var newLogs_1 = __values(newLogs), newLogs_1_1 = newLogs_1.next(); !newLogs_1_1.done; newLogs_1_1 = newLogs_1.next()) {
          var log = newLogs_1_1.value;
          var newRow = [];
          try {
            for (var LOG_HEADER_ROW_1 = (e_4 = void 0, __values(LOG_HEADER_ROW)), LOG_HEADER_ROW_1_1 = LOG_HEADER_ROW_1.next(); !LOG_HEADER_ROW_1_1.done; LOG_HEADER_ROW_1_1 = LOG_HEADER_ROW_1.next()) {
              var col = LOG_HEADER_ROW_1_1.value;
              // handle special cols first
              switch (col) {
                case "detail":
                  var reasonId = parseInt(log.reason);
                  newRow.push(REASON_MAP.get(reasonId));
                  continue;
              }
              if (log.hasOwnProperty(col)) {
                newRow.push(log[col]);
                continue;
              }
              newRow.push("");
            }
          }
          catch (e_4_1) { e_4 = { error: e_4_1 }; }
          finally {
            try {
              if (LOG_HEADER_ROW_1_1 && !LOG_HEADER_ROW_1_1.done && (_b = LOG_HEADER_ROW_1["return"])) _b.call(LOG_HEADER_ROW_1);
            }
            finally { if (e_4) throw e_4.error; }
          }
          newRows.push(newRow);
        }
      }
      catch (e_3_1) { e_3 = { error: e_3_1 }; }
      finally {
        try {
          if (newLogs_1_1 && !newLogs_1_1.done && (_a = newLogs_1["return"])) _a.call(newLogs_1);
        }
        finally { if (e_3) throw e_3.error; }
      }
    };
    while (true) {
      var response = requestApiResponse(endpoint, params);
      var entries = response.data.list;
      if (entries.length === 0) {
        // reached the end of logs
        break;
      }
      if (stopAtId) {
        var stopAtIdx = entries.findIndex(function (entry) { return entry.id === stopAtId; });
        if (stopAtIdx >= 0) {
          addLogsToNewRows(entries.slice(0, stopAtIdx));
          break;
        }
      }
      addLogsToNewRows(entries);
      params.set(API_END_ID, entries[entries.length - 1].id);
    }
    newRows.push.apply(newRows, __spreadArray([], __read(curValues.slice(1))));
    logSheet.getRange(2, 1, newRows.length, LOG_HEADER_ROW.length).setValues(newRows);

    var dashboardSheet = SpreadsheetApp.getActive().getSheetByName('Dashboard');
    dashboardSheet.getRange(bannerSettingsForImport[logSheetInfo.sheetName]['range_dashboard_length']).setValue(newRows.length);
    settingsSheet.getRange(bannerSettingsForImport[logSheetInfo.sheetName]['range_status']).setValue("Found: " + ((newRows.length + 1) - (previousRowLength)));
  }
}


var PRIMOGEM_SHEET_INFO = {
  sheetName: SHEET_NAME_PRIMOGEM_LOG,
  apiPath: "/ysulog/api/getPrimogemLog"
};

function getPrimogemLog() {
  writeLogToSheet(PRIMOGEM_SHEET_INFO);
};

var CRYSTAL_SHEET_INFO = {
  sheetName: SHEET_NAME_CRYSTAL_LOG,
  apiPath: "/ysulog/api/getCrystalLog",
}

function getCrystalLog() {
  writeLogToSheet(CRYSTAL_SHEET_INFO);
};
var bannerSettingsForImport = {
  "Primogem Log": { "range_status": "E26", "range_toggle": "E19", "range_dashboard_length": "C15" },
  "Crystal Log": { "range_status": "E27", "range_toggle": "E20", "range_dashboard_length": "C20" },
};
function importFromAPI() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  settingsSheet.getRange("E24").setValue(new Date());
  settingsSheet.getRange("E25").setValue("");

  var bannerName;
  var bannerSheet;
  var bannerSettings;
  // Clear status
  for (var i = 0; i < nameOfLogHistorys.length; i++) {
    bannerName = nameOfLogHistorys[i];
    bannerSettings = bannerSettingsForImport[bannerName];
    settingsSheet.getRange(bannerSettings['range_status']).setValue("");
  }
  for (var i = 0; i < nameOfLogHistorys.length; i++) {
    bannerName = nameOfLogHistorys[i];
    bannerSettings = bannerSettingsForImport[bannerName];
    var isToggled = settingsSheet.getRange(bannerSettings['range_toggle']).getValue();
    if (isToggled == true) {
      bannerSheet = SpreadsheetApp.getActive().getSheetByName(bannerName);
      if (bannerSheet) {
        if (bannerName == SHEET_NAME_CRYSTAL_LOG) {
          getCrystalLog();
        } else if (bannerName == SHEET_NAME_PRIMOGEM_LOG) {
          getPrimogemLog();
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
