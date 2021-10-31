// for license and source, visit https://github.com/3096/primorina

function onOpen( ){
    var ui = SpreadsheetApp.getUi();
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
    if (!settingsSheet) {
        ui.createMenu('TCBS')
        .addItem('Initialise', 'updateItemsList')
        .addToUi();
    } else {
      getDefaultMenu();
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
    }
    var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
    if (!dashboardSheet) {
      if (!sheetSource) {
        sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
      }
      var sheetDashboardSource = sheetSource.getSheetByName(SHEET_NAME_DASHBOARD);
      if (sheetDashboardSource) {
        dashboardSheet = sheetDashboardSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
        dashboardSheet.setName(SHEET_NAME_DASHBOARD);
      }
    }
    return settingsSheet;
}