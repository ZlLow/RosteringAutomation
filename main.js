/**
 * A Fixed List of Recurrence Rate
 */
const RecurrenceRate = Object.freeze({
    Daily: Symbol("Daily"),
    Monthly: Symbol("Monthly"),
    Yearly: Symbol("Yearly"),
    Once: Symbol("Once")
})

function main() {
    let masterFolder = folder.retrieveFolderByName("Master Folder");
    let eventFile = file.retrieveFileByName("Event Crew",masterFolder);
    let ss = SpreadsheetApp.open(eventFile);
    onOpenSpreadsheetTrigger(ss,"createMasterMenu");
}

function createCurrentYearHierarchy() {
    try {
        let year = dateHelper.getYear();
        var hierarchy =
            {
                parentFolder: `ESS ${year}`,
                subDirectories: ["1. Events (OTH) - Quotations x Timesheet x Invoice x Payout", "Availability", "Rostering"],
                args:
                    [{
                        parentFolder: "1. Events (OTH) - Quotations x Timesheet x Invoice x Payout",
                        subDirectories: [`OTH #${year}-999 Completed Projects ${year}`, `OTH #${year}-998 Cancelled Projects ${year}`]
                    }]
            };
        var rootFolder = folder.retrieveFolderByName("ESS Main Folder");
        folder.generateFolderHierarchy(hierarchy, rootFolder);
    } catch (e) {
        ErrorHandler.insertErrorLog(e);
    }
}

function createCurrentYearSpreadsheet() {
    try {
        let year = dateHelper.getYear();
        let month = dateHelper.getMonthName().slice(0,3);
        var rootFolder = folder.retrieveFolderByName(`ESS ${year}`);
        if (!rootFolder)
            createYearlyHierarchy();
        rootFolder = folder.retrieveFolderByName(`ESS ${year}`);
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