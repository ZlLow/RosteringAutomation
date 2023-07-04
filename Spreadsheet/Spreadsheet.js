/**
 * A List of Fixed Types to generate Spreadsheet
 */
const TemplateType = Object.freeze({
  Availability: Symbol("Availability"),
  Timesheet: Symbol("Timesheet"),
  Roster: Symbol("Roster"),
  Individuals: Symbol("Individuals")
})

/**
 * A class that creates templates and format all sheets within a particular spreadsheet
 */
class Spreadsheet {
  /**
   * Creation of class must require a spreadsheet file
   */
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  /**
   * Get Name of the Spreadsheet
   */
  getName() {
    return this.spreadsheet.getName();
  }
  /**
   * Get Sheet By Name
   * @params {string} sheetName: Name of the specific sheet
   * return {Sheet} return a customized class of Sheet
   */
  getSheet(sheetName) {
    if (typeof sheetName !== "string")
      throw new TypeError("Please choose the correct format");
    let sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return null;
    }
    return new Sheet(sheet);
  }

  /**
   * Insert Sheet Sheet
   * @params {string} sheetName: Name of the Month (Also name of the sheet)
   * @return {Sheet} sheet: Returns a New Object Sheet
   */
  insertSheet(sheetName) {
    if ((typeof sheetName !== "string"))
      throw new TypeError("Please choose the correct format!")
    let sheets = this.spreadsheet.getSheets();
    let foundSheet = sheets.find(sheet => sheet.getName() === sheetName);
    if (sheets.length === 1 &&
      sheets[0].getName() === "Sheet1") {
      console.log("Renaming Sheet Name to Month");
      foundSheet = sheets[0].setName(sheetName);
    }
    else if (!foundSheet) {
      console.log("Unable to find the correct sheet! Inserting Sheet into spreadsheet");
      foundSheet = this.spreadsheet.insertSheet(sheetName);
    }
    return new Sheet(foundSheet);
  }

  /**
   * Generats a Sheet template
   * @params {TemplateType} type: A Fixed List that detects type of template to generate
   * @params {string} sheetName: Name of the sheet
   * return {Sheet} sheet: The sheet that was targeted
   */
  generateTemplate(type = TemplateType, sheetName) {
    if ((!type))
      throw new TypeError("Please choose the correct format!")
    if (type === TemplateType.Custom && !Object.keys(args).length)
      throw new Error("Unable to process customized template as no argument is provided")
    let sheet = this.insertSheet(sheetName);
    let year = dateHelper.getYear();
    var rootFolder = folder.retrieveFolderByName(`ESS ${year}`)
    switch (type) {
      case TemplateType.Availability:
        sheet.createAvailabilityTemplate(sheetName);
        sheet.insertAvailabilityData("Individuals", "", sheetName);
        SpreadsheetApp.setActiveSpreadsheet(this.spreadsheet);
        onOpenSpreadsheetTrigger(this.spreadsheet, "createAvailMenu");
        onOpenSpreadsheetTrigger(this.spreadsheet, "createMonthlyAvailSpreadsheet");
        break;
      case TemplateType.Roster:
        sheet.createRosterTemplate(sheetName);
        let masterFile = file.retrieveFileByName("Event Crew");
        let masterSS = new Spreadsheet(SpreadsheetApp.open(masterFile));
        let masterSheet = masterSS.getSheet("Sheet1");
        sheet.insertRosterData("Yearly Availability", masterSheet, rootFolder, sheetName);
        SpreadsheetApp.setActiveSpreadsheet(this.spreadsheet);
        onOpenSpreadsheetTrigger(this.spreadsheet, "createMonthlyRosterSpreadsheet");
        onOpenSpreadsheetTrigger(this.spreadsheet, "createRosterMenu");
        break;
      case TemplateType.Timesheet:
        sheet.createTimesheetTemplate();
        break;
      case TemplateType.Individuals:
        sheet.createIndividualTemplate(sheetName);
        break;
      default:
        throw new SyntaxError("There is an issue taking in the parameter: (type)!");
    }
    return sheet;
  }

  /**
   * Insert Data to Sheet
   * @params {TemplateType} type: A Fixed List that detects type of template to generate
   * @params {string} sheetName: Name of the Sheet
   * @params {Array of JSON} args: Arguments for creation of Timesheet Data
   * Data Sample:
   * [
   *           {
   *            id: 1,
   *            name: "David",
   *            role: ["IC","Usher"],
   *            date: ["1 June","2 June"]
   *           },
   *           {
   *            id: 3,
   *            name: "Aaron",
   *            role: ["IC"],
   *            date: ["1 June"]
   *           },
   *]
   * @returns {Sheet} sheet : A new Object of Sheet;
   */
  insertDataToSheet(type = TemplateType, sheetName, args = []) {
    if ((!type) && typeof sheetName !== "string")
      throw new TypeError("Please choose the correct format!")
    let foundSheet = this.spreadsheet.getSheets().find(sheet => sheet.getName() === sheetName);
    if (!foundSheet)
      throw new Error("Unable to find the sheet Name. Please ensure that the sheet name is correct!");
    let sheet = new Sheet(foundSheet);
    let year = dateHelper.getYear();
    var rootFolder = folder.retrieveFolderByName(`ESS ${year}`)
    switch (type) {
      case TemplateType.Availability:
        sheet.insertAvailabilityData("Individuals", "", sheetName);
        break;
      case TemplateType.Roster:
        let masterFile = file.retrieveFileByName("Event Crew");
        let masterSS = new Spreadsheet(SpreadsheetApp.open(masterFile));
        let masterSheet = masterSS.getSheet("Sheet1");
        sheet.insertRosterData("Yearly Availability", masterSheet, rootFolder, sheetName);
        break;
      case TemplateType.Timesheet:
        sheet.insertTimesheetData(args);
        break;
      default:
        throw new SyntaxError(`There is an issue taking in the parameter: ${type}!`);
    }
    return sheet;


  }
}