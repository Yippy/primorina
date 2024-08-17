// for license and source, visit https://github.com/3096/primorina

var dashboardEditRange = [
  "I5", // Status cell
  "AP40", // Document Version
  "AT53", // Current Document Version
  "AH41", // Document Status
  "AV1", // Name of drop down 1 (import)
  "AV2", // Name of drop down 2 (auto import)
  "AG14", // Selection
  "", // URL [NOT NEEDED PROMPT HANDLES URL]
  "AV3", // Name of drop down 3 (HoYoLAB)
  "AG16", // Name of user input
  "A53", // Dashboard embedded or add-on notification
];

// Cells that needs refreshing
var dashboardRefreshRange = [
];

function onEdit(e) {
  // When amending the language, the start of the week refer to the selection first day
  var ss = e.range.getSheet();
  if (e.range.getA1Notation() === 'B2' && ss.getName() === SHEET_NAME_SETTINGS) {
    ss.getRange("B5").setValue(ss.getRange("J3").getValue());

    for (var i = 0; i < MONTHLY_SHEET_NAME.length; i++) {
      var monthlySheet = SpreadsheetApp.getActive().getSheetByName(MONTHLY_SHEET_NAME[i]);
      refreshMonthlyMonthText(monthlySheet,ss);
    }
  }
}

function refreshMonthlyMonthText(monthlySheet, settingsSheet) {
  if (monthlySheet && settingsSheet) {
    var restoreRanges = monthlySheet.getRange("A1").getValue();
    restoreRanges = String(restoreRanges).split(",");
    if (restoreRanges.length == 2) {
      var monthIndex = Number(monthlySheet.getRange(restoreRanges[0]).getValue());
      var monthNameInSelection = settingsSheet.getRange(2,17+monthIndex).getValue();
      monthlySheet.getRange(restoreRanges[1]).setValue(monthNameInSelection);
    }
  }
}

function loadHoYoLabData(settingsSheet) {
  if (settingsSheet.getRange("D30").isPartOfMerge()) {
    // Migration process for v2.0
    settingsSheet.getRange("D30:E30").breakApart();
    settingsSheet.getRange("D30").setValue("cookie_token_v2").setBackground("#cccccc");
    settingsSheet.getRange("E30").setValue(settingsSheet.getRange("D31").getValue()).setBackground("white").setNumberFormat("@");

    settingsSheet.getRange("D31:E31").breakApart();
    settingsSheet.getRange("D31").setValue("account_id_v2").setBackground("#cccccc");
    settingsSheet.getRange("E31").setValue(settingsSheet.getRange("D33").getValue()).setBackground("white").setNumberFormat("0");

    settingsSheet.getRange("D32:E32").breakApart();
    settingsSheet.getRange("D32").setValue("account_mid_v2").setBackground("#cccccc");
    settingsSheet.getRange("E32").setValue("").setBackground("white").setNumberFormat("@");

    settingsSheet.getRange("D33:E33").breakApart();
    settingsSheet.getRange("D33").setValue("Genshin Impact UID").setBackground("#cccccc");
    settingsSheet.getRange("E33").setValue("").setBackground("white").setNumberFormat("0");
  }
  ltokenInput = settingsSheet.getRange("E30").getValue();

  ltuidInput = settingsSheet.getRange("E31").getValue();
  if(ltuidInput) {
    ltuidInput.toFixed(0);
  }
  ltmidInput = settingsSheet.getRange("E32").getValue();
  guidInput = settingsSheet.getRange("E33").getValue();
  if(guidInput) {
    guidInput.toFixed(0);
  }
}

function importButtonScript() {
  var settingsSheet = getSettingsSheet();
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
  if (dashboardSheet && settingsSheet) {
    var userImportSelection = dashboardSheet.getRange(dashboardEditRange[4]).getValue();
    var userAutoImportSelection = dashboardSheet.getRange(dashboardEditRange[5]).getValue();
    var importSelectionText = dashboardSheet.getRange(dashboardEditRange[6]).getValue();
    var userInputText = dashboardSheet.getRange(dashboardEditRange[9]).getValue();

    // Check if the selection is not Import from old Document
    if (userImportSelection != importSelectionText) {
      if (userAutoImportSelection == importSelectionText) {
        userInputText += "\n\nNote\nEntering an empty URL will try to load from previously saved Settings";
        // Too long to display URL to user for Auto Import from miHoYo
      } else {
        loadHoYoLabData(settingsSheet);
        userInputText = "Note\nEntering an empty HoYoLAB ltoken/cookie_token_v2 will try to load from previously saved Settings";
        userInputText += "\n\nPreviously Saved HoYoLab Cookies:\ncookie_token_v2: "+ltokenInput;
        userInputText += "\nltuid/account_id_v2: "+ltuidInput;
        userInputText += "\naccount_mid_v2: "+ltmidInput;
        userInputText += "\nGenshin Impact UID: "+guidInput;
        
        userInputText += "\n\nEnter new cookie_token_v2 in textfield:";
      }
    }

    const resultURL = displayUserPrompt(importSelectionText, userInputText, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    var button = resultURL.getSelectedButton();
    if (button == SpreadsheetApp.getUi().Button.OK) {
      var urlInput = resultURL.getResponseText();
      if (urlInput.length > 0) {
        if (userImportSelection == importSelectionText) {
          settingsSheet.getRange("D6").setValue(urlInput);
          importDataManagement();
        } else if (userAutoImportSelection == importSelectionText) {
          settingsSheet.getRange("D17").setValue(urlInput);
          importFromAPI();
        } else {
          settingsSheet.getRange("E30").setValue(urlInput);
          ltokenInput = urlInput;
          loadHoYoLABUIDPopup(settingsSheet);
        }
      } else {
        if (userImportSelection == importSelectionText) {
          SpreadsheetApp.getActiveSpreadsheet().toast("Error URL provided is empty, stopping import function.", importSelectionText);
        } else {
          var loadPreviousSetting = "";
          if (userAutoImportSelection == importSelectionText) {
            loadPreviousSetting = settingsSheet.getRange("D17").getValue();
          } else {
            loadPreviousSetting = ltokenInput;
          }
          if (loadPreviousSetting.length > 0) {
            var userInputText = 'The user input is empty,\nwould you like to reuse the previously stored data from settings?';

            if (userAutoImportSelection == importSelectionText) {
              // Too long to display URL to user for Auto Import from miHoYo
            } else {
              // Friendly reminder to user of HoYoLab detailed saving in settings
              userInputText = 'Would you like to reuse these previously stored data from settings?';
              userInputText += "\n\nHoYoLab cookie_token_v2: "+loadPreviousSetting;
              userInputText += "\nHoYoLAB account_id_v2: "+ltuidInput;
              userInputText += "\nHoYoLAB account_mid_v2: "+ltmidInput;
              userInputText += "\nGenshin Impact UID: "+guidInput;
            }
            const result = displayUserAlert(importSelectionText, userInputText);
            if (result == SpreadsheetApp.getUi().Button.OK) {
              // User wants to reuse previously stored data
              if (userAutoImportSelection == importSelectionText) {
                importFromAPI();
              } else {
                importFromHoYoLAB();
              }
            }
          } else {
            // There is no previous settings, prompt user an error message
            SpreadsheetApp.getActiveSpreadsheet().toast("Error previous Settings is empty, stopping import function.", importSelectionText);
          }
        }
      }
    }
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Unable to find 'Dashboard' or 'Settings'", "Missing Sheets");
  }
}

function displayReadme() {
  var sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
  if (sheetSource) {
    // Avoid Exception: You can't remove all the sheets in a document.Details
    var placeHolderSheet = null;
    if (SpreadsheetApp.getActive().getSheets().length == 1) {
      placeHolderSheet = SpreadsheetApp.getActive().insertSheet();
    }
    var sheetToRemove = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_README);
    if (sheetToRemove) {
      // If exist remove from spreadsheet
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToRemove);
    }
    var sheetREADMESource;

    // Add Language
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
    if (settingsSheet) {
      var languageFound = settingsSheet.getRange(2, 2).getValue();
      sheetREADMESource = sheetSource.getSheetByName(SHEET_NAME_README + "-" + languageFound);
    }
    if (sheetREADMESource) {
      // Found language
    } else {
      // Default
      sheetREADMESource = sheetSource.getSheetByName(SHEET_NAME_README);
    }

    sheetREADMESource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(SHEET_NAME_README);

    // Remove placeholder if available
    if (placeHolderSheet) {
      // If exist remove from spreadsheet
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(placeHolderSheet);
    }
    var sheetREADME = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_README);
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

function moveToSheetByName(nameOfSheet: string) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(nameOfSheet);
  if (sheet) {
    sheet.activate();
  } else {
    var findSheetFromSource;
    var sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
    if (sheetSource) {
      findSheetFromSource = sheetSource.getSheetByName(nameOfSheet);
      if (findSheetFromSource) {
        sheet = findSheetFromSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(nameOfSheet);
      }
    }
    if (sheet) {
      sheet.activate();
      const title = "Found";
      const message = "'" + nameOfSheet + "' was copied from Source.";
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    } else {
      const title = "Error";
      const message = "Unable to find sheet named '" + nameOfSheet + "' and source is unavailable.";
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    }
  }
}

const moveToDashboardSheet = () => moveToSheetByName(SHEET_NAME_DASHBOARD);
const moveToSettingsSheet = () => moveToSheetByName(SHEET_NAME_SETTINGS);
const moveToChangelogSheet = () => moveToSheetByName(SHEET_NAME_CHANGELOG);
const moveToReadmeSheet = () => moveToSheetByName(SHEET_NAME_README);
const moveToPrimogemLogSheet = () => moveToSheetByName(SHEET_NAME_PRIMOGEM_LOG);
const moveToPrimogemYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_PRIMOGEM_YEARLY_REPORT);
const moveToPrimogemMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_PRIMOGEM_MONTHLY_REPORT);
const moveToCrystalLogSheet = () => moveToSheetByName(SHEET_NAME_CRYSTAL_LOG);
const moveToCrystalYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_CRYSTAL_YEARLY_REPORT);
const moveToCrystalMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_CRYSTAL_MONTHLY_REPORT);
const moveToResinLogSheet = () => moveToSheetByName(SHEET_NAME_RESIN_LOG);
const moveToResinYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_RESIN_YEARLY_REPORT);
const moveToResinMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_RESIN_MONTHLY_REPORT);
const moveToArtifactLogSheet = () => moveToSheetByName(SHEET_NAME_ARTIFACT_LOG);
const moveToArtifactYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_ARTIFACT_YEARLY_REPORT);
const moveToArtifactMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_ARTIFACT_MONTHLY_REPORT);
const moveToMoraLogSheet = () => moveToSheetByName(SHEET_NAME_MORA_LOG);
const moveToMoraYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_MORA_YEARLY_REPORT);
const moveToMoraMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_MORA_MONTHLY_REPORT);
const moveToWeaponLogSheet = () => moveToSheetByName(SHEET_NAME_WEAPON_LOG);
const moveToWeaponYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_WEAPON_YEARLY_REPORT);
const moveToWeaponMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_WEAPON_MONTHLY_REPORT);
const moveToStarglitterLogSheet = () => moveToSheetByName(SHEET_NAME_STARGLITTER_LOG);
const moveToStardustLogSheet = () => moveToSheetByName(SHEET_NAME_STARDUST_LOG);
const moveToMasterlessYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_MASTERLESS_YEARLY_REPORT);
const moveToMasterlessMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_MASTERLESS_MONTHLY_REPORT);
const moveToKeyItemsSheet = () => moveToSheetByName(SHEET_NAME_KEY_ITEMS);