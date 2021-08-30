// for license and source, visit https://github.com/3096/primorina

function reorderSheets() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
  if (settingsSheet) {
    var sheetsToSort = settingsSheet.getRange(11, 2, 11, 1).getValues();
    Logger.log(sheetsToSort);

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
          var sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
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
        for (var i = 0; i < NAME_OF_LOG_HISTORIES.length; i++) {
          var bannerImportSheet = importSource.getSheetByName(NAME_OF_LOG_HISTORIES[i]);

          var numberOfRows = bannerImportSheet.getMaxRows() - 1;
          var range = bannerImportSheet.getRange(2, 1, numberOfRows, 5);

          if (bannerImportSheet && numberOfRows > 0) {
            var bannerSheet = SpreadsheetApp.getActive().getSheetByName(NAME_OF_LOG_HISTORIES[i]);

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
  var sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
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
        var settingOptionNum = parseInt(settingsOptionRanges[i]);

        var sheetAvailableSelectionSource = sheetSource.getSheetByName(nameOfBanner);
        var storedSheet;
        if (isSkipString == "YES") {
          // skip - disabled by source
        } else {
          if (sheetAvailableSelectionSource) {
            if (settingOptionNum) {
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
