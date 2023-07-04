/* -----------------------------------------------------------------------------------Time Based Trigger --------------------------------------------------------------------------------------------*/

/**
 * Create Yearly Folder Hierarchy
 * Automatically be called by Trigger Yearly.
 * Constraints: 1. Cannot contain any parameters
 *              2. Any Changes to the code will cause changes within the yearly trigger
 */
function createYearlyHierarchy() {
  try {
    let nextYear = dateHelper.getNextYear().getFullYear().toString();
    var hierarchy =
    {
      parentFolder: `ESS ${nextYear}`,
      subDirectories: ["1. Events (OTH) - Quotations x Timesheet x Invoice x Payout", "Availability", "Rostering"],
      args:
        [{
          parentFolder: "1. Events (OTH) - Quotations x Timesheet x Invoice x Payout",
          subDirectories: [`OTH #${nextYear}-999 Completed Projects ${nextYear}`, `OTH #${nextYear}-998 Cancelled Projects ${nextYear}`]
        }]
    };
    var rootFolder = folder.retrieveFolderByName("ESS Main Folder");
    folder.generateFolderHierarchy(hierarchy, rootFolder);
  } catch (e) {
    ErrorHandler.insertErrorLog(e);
  }
}

/**
 * Create Yearly Spreadsheet
 * Automatically be called by Trigger Yearly
 * Constraints:  1. Cannot contain any parameters
 *               2. Any changes to the code will cause changes within yearly trigger
 */
function createYearlySpreadsheets() {
  try {
    let nextYear = dateHelper.getNextYear();
    let month = dateHelper.getMonthName(nextYear).slice(0, 3);
    var rootFolder = folder.retrieveFolderByName(`ESS ${nextYear.getFullYear().toString()}`);
    if (!rootFolder)
      createYearlyHierarchy();
    rootFolder = folder.retrieveFolderByName(`ESS ${nextYear.getFullYear().toString()}`);
    let availFolder = folder.retrieveFolderByName("Availability", rootFolder);
    let rosterFolder = folder.retrieveFolderByName("Rostering", rootFolder);
    let availFile = file.retrieveFileByName("Yearly Availability", availFolder);
    let rosterFile = file.retrieveFileByName("Yearly Roster Mastersheet", rosterFolder);
    let availSS = !availFile ? file.createSpreadsheetToFolder("Yearly Availability", availFolder) : SpreadsheetApp.open(availFile);
    let rosterSS = !rosterFile ? file.createSpreadsheetToFolder("Yearly Roster Mastersheet", rosterFolder) : SpreadsheetApp.open(rosterFile);
    let avail = new Spreadsheet(availSS);
    let roster = new Spreadsheet(rosterSS);
    avail.generateTemplate(TemplateType.Availability, month);
    roster.generateTemplate(TemplateType.Roster, month);
  } catch (e) {
    ErrorHandler.insertErrorLog(e);
  }
}

/*--------------------------------------------------------------------------Event Based Trigger ---------------------------------------------------------------------------------------------------- */

/**
 * Create Monthly Spreadsheet
 * @params {Event} spreadsheetEvent: An Event which contains the spreadsheet that is being targeted
 */
function createMonthlyAvailSpreadsheet(spreadsheetEvent) {
  try {
    let source = spreadsheetEvent.source;
    var triggers = ScriptApp.getProjectTriggers();
    let triggerID = spreadsheetEvent.triggerUid;
    let { date, month } = dateHelper.getCurrentDates();
    if (triggers.find(trigger => trigger.getUniqueId() === triggerID)) {
      if (month === 11)
        ScriptApp.deleteTrigger(this);
      else {
        let ss = new Spreadsheet(source);
        let nextMonth = dateHelper.getNextMonth().slice(0, 3);
        if (ss.getSheet(nextMonth))
          return;
        if (date > 10) {
          ss.generateTemplate(TemplateType.Availability, nextMonth);
        }
      }
    }
  } catch (e) {
    let ui = SpreadsheetApp.getUi();
    let handler = new ErrorHandler(ui);
    handler.createAlert(e);
  }
}

/**
 * Create Monthly Spreadsheet
 * @params {Event} spreadsheetEvent: An Event which contains the spreadsheet that is being targeted
 */
function createMonthlyRosterSpreadsheet(spreadsheetEvent) {
  try {
    let source = spreadsheetEvent.source;
    var triggers = ScriptApp.getProjectTriggers();
    let triggerID = spreadsheetEvent.triggerUid;
    let { date, month } = dateHelper.getCurrentDates();
    if (triggers.find(trigger => trigger.getUniqueId() === triggerID)) {
      if (month === 11)
        ScriptApp.deleteTrigger(this);
      else {
        let ss = new Spreadsheet(source);
        let nextMonth = dateHelper.getNextMonth().slice(0, 3);
        if (ss.getSheet(nextMonth))
          return;
        if (date > 10) {
          ss.generateTemplate(TemplateType.Roster, nextMonth);
        }
      }
    }
  } catch (e) {
    let ui = SpreadsheetApp.getUi();
    let handler = new ErrorHandler(ui);
    handler.createAlert(e);
  }
}


/**
 * Trigger Function which create Menu. This UI allows easier update based on visuals
 * @params {Event} spreadsheetEvent: An Event which contains the spreadsheet that is being targeted
 */
function createAvailMenu(spreadsheetEvent) {
  try {
    let ss = spreadsheetEvent.source;
    Menu.createMenu(ss, "Avail Menu", ["Update Data", "refreshAvailabilityData"]);
  } catch (e) {
    let ui = SpreadsheetApp.getUi();
    let handler = new ErrorHandler(ui);
    handler.createAlert(e);
  }
}

/**
 * Trigger Function which create Menu. This UI allows easier update based on visuals
 * Menus are customized based on the Spreadsheet
 * @params {Event} spreadsheetEvent: An Event which contains the spreadsheet that is being targeted
 */
function createRosterMenu(spreadsheetEvent) {
  try {
    let ss = spreadsheetEvent.source;
    Menu.createMenu(ss, "Roster Menu", ["Update Data", "updateRosterSpreadsheet"], ["Generate Timesheet Data", "generateTimesheet"]);
  } catch (e) {
    let ui = SpreadsheetApp.getUi();
    let handler = new ErrorHandler(ui);
    handler.createAlert(e);
  }
}

/**
 * Trigger Function which create Menu. This UI allows easier update based on visuals
 * Menus are customized based on the Spreadsheet
 * @params {Event} spreadsheetEvent: An Event which contains the spreadsheet that is being targeted
 */
function createMasterMenu(spreadsheetEvent) {
  try {
    let ss = spreadsheetEvent.source;
    Menu.createMenu(ss, "Master Menu", ["Generate Availability Template", "generateIndividualSpreadsheet"]);
  } catch (e) {
    let ui = SpreadsheetApp.getUi();
    let handler = new ErrorHandler(ui);
    handler.createAlert(e);
  }
}

