// for license and source, visit https://github.com/3096/primorina
function onInstall(e) {
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    generateInitialiseToolbar();
  } else {
    onOpen(e);
  }
}

function onOpen(e) {
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    generateInitialiseToolbar();
  } else {
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
    if (!settingsSheet) {
      generateInitialiseToolbar();
    } else {
      getDefaultMenu();
    }
    checkLocaleIsSetCorrectly();
  }
}

function generateInitialiseToolbar() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('TCBS')
  .addItem('Initialise', 'updateItemsList')
  .addToUi();
}

function displayUserPrompt(titlePrompt: string, messagePrompt: string) {
  const ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    titlePrompt,
    messagePrompt,
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  return result;
}

function displayUserAlert(titleAlert: string, messageAlert: string) {
  const ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    titleAlert,
    messageAlert,
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  return result;
}

/* Ensure Sheets is set to the supported locale due to source document formula */
function checkLocaleIsSetCorrectly() {
  var currentLocale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  if (currentLocale != SHEET_SOURCE_SUPPORTED_LOCALE) {
    SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetLocale(SHEET_SOURCE_SUPPORTED_LOCALE);
    var message = 'To ensure compatibility with formula from source document, your locale "'+currentLocale+'" has been changed to "'+SHEET_SOURCE_SUPPORTED_LOCALE+'"';
    var title = 'Sheets Locale Changed';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function getDefaultMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('TCBS')
  .addSeparator()
  .addSubMenu(ui.createMenu('Data Management')
          .addItem('Import', 'importDataManagement')
          .addSeparator()
          .addItem('Auto Import from miHoYo', 'importFromAPI')
          .addItem('Auto Import from HoYoLAB', 'importFromHoYoLAB')
          )
  .addSeparator()
  .addItem('Quick Update', 'quickUpdate')
  .addItem('Update Items', 'updateItemsList')
  .addItem('Get Latest README', 'displayReadme')
  .addToUi();
}
// set data validation
var listOfSheets = [
  SHEET_NAME_DASHBOARD,
  SHEET_NAME_README,
  SHEET_NAME_CHANGELOG,
  SHEET_NAME_SETTINGS,
  SHEET_NAME_PRIMOGEM_MONTHLY_REPORT,
  SHEET_NAME_PRIMOGEM_YEARLY_REPORT,
  SHEET_NAME_PRIMOGEM_LOG,
  SHEET_NAME_CRYSTAL_MONTHLY_REPORT,
  SHEET_NAME_CRYSTAL_YEARLY_REPORT,
  SHEET_NAME_CRYSTAL_LOG,
  SHEET_NAME_RESIN_MONTHLY_REPORT,
  SHEET_NAME_RESIN_YEARLY_REPORT,
  SHEET_NAME_RESIN_LOG,
  SHEET_NAME_MORA_MONTHLY_REPORT,
  SHEET_NAME_MORA_YEARLY_REPORT,
  SHEET_NAME_MORA_LOG,
  SHEET_NAME_ARTIFACT_MONTHLY_REPORT,
  SHEET_NAME_ARTIFACT_YEARLY_REPORT,
  SHEET_NAME_ARTIFACT_LOG,
  SHEET_NAME_ARTIFACT_ITEMS,
  SHEET_NAME_WEAPON_LOG,
  SHEET_NAME_WEAPON_MONTHLY_REPORT,
  SHEET_NAME_WEAPON_YEARLY_REPORT,
  SHEET_NAME_KEY_ITEMS
];
function getSettingsSheet() {
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
    var sheetSource;
    if (!settingsSheet) {
      sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
      var sheetSettingSource = sheetSource.getSheetByName(SHEET_NAME_SETTINGS);
      if (sheetSettingSource) {
        settingsSheet = sheetSettingSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
        settingsSheet.setName(SHEET_NAME_SETTINGS);
        getDefaultMenu();
      }
    } else {
      settingsSheet.getRange("H1").setValue(SCRIPT_VERSION);
      /* Check migration for settings required */
      // REMOVE FOR V2.0
      var bannerSettings = LOG_RANGES[SHEET_NAME_ARTIFACT_LOG];
      var isAvailable = settingsSheet.getRange("A31").getValue();
      if(isAvailable == ""){
        // Migration step for 1.10, missing Artifact toggle
        settingsSheet.getRange(13,4,1,2).setBorder(false, false, false, false, false, false).breakApart();
        settingsSheet.getRange("D13").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_MORA_LOG);
        settingsSheet.getRange("E13").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setValue(settingsSheet.getRange("E12").getValue());
        settingsSheet.getRange("D12").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_ARTIFACT_LOG);
        settingsSheet.getRange("E12").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setValue('NOT DONE');
        settingsSheet.getRange(14,4,2,2).breakApart();

        settingsSheet.getRange(14,4,1,2).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID).mergeAcross();
        settingsSheet.getRange("D14").setFontColor("black").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setValue('Auto Import');
        settingsSheet.getRange(15,4,1,2).mergeAcross();
        settingsSheet.getRange("D15").setFontColor("red").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue('Using Genshin Impact Feedback URL to call API with AUTH_KEY');

        settingsSheet.getRange("D22").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_ARTIFACT_LOG);
        settingsSheet.getRange("E22").setFontSize(10).setBackground("white").insertCheckboxes().setValue(true);
        settingsSheet.getRange("D29").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_ARTIFACT_LOG);

        settingsSheet.getRange("A28").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(18);
        settingsSheet.getRange("A29").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(19);
        settingsSheet.getRange("A30").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(20);
        settingsSheet.getRange("A31").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(21);
        settingsSheet.getRange("A32").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(22);
        settingsSheet.getRange("A33").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(23);
        
        settingsSheet.getRange("B27").setValue(SHEET_NAME_ARTIFACT_MONTHLY_REPORT);
        settingsSheet.getRange("B28").setValue(SHEET_NAME_ARTIFACT_YEARLY_REPORT);
        settingsSheet.getRange("B29").setValue(SHEET_NAME_ARTIFACT_LOG);
        settingsSheet.getRange("B30").setValue(SHEET_NAME_ARTIFACT_ITEMS);
        settingsSheet.getRange("B31").setValue(SHEET_NAME_KEY_ITEMS);

        // Load Artifact Log Sheet if missing
        var artifactLogSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ARTIFACT_LOG);
        if (!artifactLogSheet) {
          if (!sheetSource) {
            sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
          }
          var sheetArtifactLogSource = sheetSource.getSheetByName(SHEET_NAME_ARTIFACT_LOG);
          artifactLogSheet = sheetArtifactLogSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
          artifactLogSheet.setName(SHEET_NAME_ARTIFACT_LOG);
        }
      }
      // Migration step for v1.12 loading Weapon user preferences
      isAvailable = settingsSheet.getRange("D43").getValue();
      if (isAvailable == "") {
        // Data Validation List updated
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(listOfSheets, true).build();
        settingsSheet.getRange(11,2,23,1).setBackground("white").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setDataValidation(rule);
        settingsSheet.getRange("B32").setValue(SHEET_NAME_WEAPON_LOG);
        settingsSheet.getRange("B33").setValue(SHEET_NAME_WEAPON_YEARLY_REPORT);

        settingsSheet.getRange(14,4,1,2).setBorder(false, false, false, false, false, false).breakApart();
        settingsSheet.getRange("D14").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_MORA_LOG);
        settingsSheet.getRange("E14").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setValue(settingsSheet.getRange("E13").getValue());
        settingsSheet.getRange("D13").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_WEAPON_LOG);
        settingsSheet.getRange("E13").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setValue('NOT DONE');
        settingsSheet.getRange(15,4,1,2).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID).mergeAcross().setFontColor("black").setFontSize(11).setFontColor("black").setFontWeight("bold").setHorizontalAlignment("center").setValue('Auto Import');

        settingsSheet.getRange(43,4,1,2).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID).mergeAcross().setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setValue("Auto Import - Cont.");
        settingsSheet.getRange(44,4,1,2).mergeAcross().setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue("Select Log to Update");
        settingsSheet.getRange(45,4,2,1).mergeAcross().setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_WEAPON_LOG);
        settingsSheet.getRange("E45").setFontSize(10).setBackground("white").insertCheckboxes().setValue(true);
        // Load Weapon Log Sheet if missing
        var weaponLogSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_WEAPON_LOG);
        if (!weaponLogSheet) {
          if (!sheetSource) {
            sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
          }
          var sheetWeaponLogSource = sheetSource.getSheetByName(SHEET_NAME_WEAPON_LOG);
          weaponLogSheet = sheetWeaponLogSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
          weaponLogSheet.setName(SHEET_NAME_WEAPON_LOG);
        }
        // Remove old Dashboard if exist
        var removeDashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
        if (removeDashboardSheet) {
          SpreadsheetApp.getActiveSpreadsheet().deleteSheet(removeDashboardSheet);
        }
        checkUserPreferenceExist(settingsSheet);
      }
    }
    var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
    var isDashboardUpToDate = false;
    // Migration step for v1.12 dashboard - double check if dashboard is old when Settings page was updated before release
    if (dashboardSheet) {
      if (dashboardSheet.getRange("C38").getValue() == SHEET_NAME_WEAPON_LOG) {
        isDashboardUpToDate = true;
      } else {
        // Remove outdated dashboard
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(dashboardSheet);
      }
    }
    if (!isDashboardUpToDate) {
      if (!sheetSource) {
        sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
      }
      var sheetDashboardSource = sheetSource.getSheetByName(SHEET_NAME_DASHBOARD);
      if (sheetDashboardSource) {
        dashboardSheet = sheetDashboardSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
        dashboardSheet.setName(SHEET_NAME_DASHBOARD);
        updateDashboard(dashboardSheet);
      }
    }
    if (SHEET_SCRIPT_IS_ADD_ON) {
      dashboardSheet.getRange(dashboardEditRange[10]).setFontColor("green").setFontWeight("bold").setHorizontalAlignment("left").setValue("Add-On Enabled");
    } else {
      dashboardSheet.getRange(dashboardEditRange[10]).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("left").setValue("Embedded Script");
    }

    return settingsSheet;
}
// REMOVE FOR V2.0
// Due to newer script, migration must be placed on User Preferences for Monthly and Yearly Report
function checkUserPreferenceExist(settingsSheet) {
  // Migration step for v1.11
  var listOfPreferences = ["NO","YES"];
  if(settingsSheet.getRange("A34").getValue() != SHEET_NAME_PRIMOGEM_LOG) {
    // Missing user preference
    settingsSheet.insertRowsAfter(39,10);

    var rulePreferences = SpreadsheetApp.newDataValidation().requireValueInList(listOfPreferences, true).build();
    var rowIndexLoop = 0;
    for (const key in userPreferences) {
      settingsSheet.getRange(34 + (3 * rowIndexLoop),1,1,2).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID).mergeAcross().setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setValue(key);
      settingsSheet.getRange(35 + (3 * rowIndexLoop),1).setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(USER_PREFERENCE_MONTHLY_REPORT);
      settingsSheet.getRange(35 + (3 * rowIndexLoop),2).setBackground("white").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setDataValidation(rulePreferences).setValue("YES");
      settingsSheet.getRange(36 + (3 * rowIndexLoop),1).setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(USER_PREFERENCE_YEARLY_REPORT);
      settingsSheet.getRange(36 + (3 * rowIndexLoop),2).setBackground("white").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setDataValidation(rulePreferences).setValue("YES");
      rowIndexLoop++;
    }
    settingsSheet.getRange(34 + (3 * rowIndexLoop),1,1,2).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID).mergeAcross().setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setValue("");
  }
  // Migration step for v1.12
  if(settingsSheet.getRange("A49").getValue() != SHEET_NAME_WEAPON_LOG) {
    // Missing Weapon user preference
    settingsSheet.insertRowsAfter(49,3);

    var rulePreferences = SpreadsheetApp.newDataValidation().requireValueInList(listOfPreferences, true).build();
    settingsSheet.getRange(49,1,1,2).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID).mergeAcross().setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_WEAPON_LOG);
    settingsSheet.getRange(50,1).setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(USER_PREFERENCE_MONTHLY_REPORT);
    settingsSheet.getRange(50,2).setBackground("white").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setDataValidation(rulePreferences).setValue("YES");
    settingsSheet.getRange(51,1).setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(USER_PREFERENCE_YEARLY_REPORT);
    settingsSheet.getRange(51,2).setBackground("white").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setDataValidation(rulePreferences).setValue("YES");
    settingsSheet.getRange(52,1,1,2).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID).mergeAcross().setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setValue("");
  }
}

function updateDashboard(dashboardSheet) {
  // Go through the available logs sheet list
  const availableSheets = NAME_OF_LOG_HISTORIES.concat(NAME_OF_LOG_HISTORIES_HOYOLAB);
  for (var i = 0; i < availableSheets.length; i++) {
    var logSheet = SpreadsheetApp.getActive().getSheetByName(availableSheets[i]);

    if (logSheet) {
      var iLastRow = logSheet.getRange(2, 2, logSheet.getLastRow(), 1).getValues().filter(String).length;
      dashboardSheet.getRange(LOG_RANGES[availableSheets[i]]['range_dashboard_length']).setValue(iLastRow);
    }
  }
}