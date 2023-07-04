/*------------------------------------------------------------------------------ Menu Trigger ------------------------------------------------------------------------------------------------------*/
/**
 * Trigger Function which updates Roster Spreadsheet
 * Insertion of Data
 */
function updateRosterSpreadsheet() {
    try {
        let activeSS = SpreadsheetApp.getActive();
        let activeSheet = activeSS.getActiveSheet();
        let monthName = activeSheet.getName();
        let ss = new Spreadsheet(activeSS);

        ss.insertDataToSheet(TemplateType.Roster, monthName);
    } catch (e) {
        console.log("There is an issue:" + e);
    }
}

/**
 * Trigger Function which refreshes Availability Data
 */
function refreshAvailabilityData() {
    try {
        let activeSS = SpreadsheetApp.getActive();
        let activeSheet = activeSS.getActiveSheet();
        let monthName = activeSheet.getName();
        let ss = new Spreadsheet(activeSS);
        ss.insertDataToSheet(TemplateType.Availability, monthName);
    } catch (e) {
        console.log("There is an issue:" + e);
    }
}

/**
 * Generate Timesheet
 * Servicing of Roster Mastersheet Data into usable Data
 * Get Availability, Roster Role, Event ID, ESS ID & Name
 */
function generateTimesheet() {
    let activeSS = SpreadsheetApp.getActive();
    let activeSheet = activeSS.getActiveSheet();
    let sheet = new Sheet(activeSheet);
    let data = sheet.getAllData();


    // Getting ALL Dates
    let dateRange = data[0].filter(value => value !== "");
    dateRange.shift();

    // Getting All Subheaders
    let eventIDIndexes = sheet.getHeader("Event ID");
    let rosteredRoleIndexes = sheet.getHeader("Rostered Role");
    let availabilityIndexes = sheet.getHeader("Availability");
    let idIndexes = sheet.getHeader("ESS ID");
    let nameIndexes = sheet.getHeader("Name");
    let firstRow = eventIDIndexes[0].rowIndex;
    let values = data.slice(firstRow + 1);

    // Retrieving All data and Formatting
    /**
     * Data Sample:
     * [{
     *   id: "OTH #2023-001"
     *   crew : [
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
     *          ]
     *         }
     * }
     * ]
     */
    let events = [];
    dateRange.forEach((value, index) => {
        let date = dateHelper.getDate(value);
        values.forEach(valueRow => {
            let avail = valueRow[availabilityIndexes[0].colIndexes[index]];
            if (avail === "Not Available" || avail === "")
                return;
            let eID = valueRow[eventIDIndexes[0].colIndexes[index]];
            let id = valueRow[idIndexes[0].colIndexes[0]];
            let name = valueRow[nameIndexes[0].colIndexes[0]];
            let role = valueRow[rosteredRoleIndexes[0].colIndexes[index]];
            var event = events.find(data => data.id == eID);
            if (events.includes(event)) {
                let crew = event.crew.find(data => data.id === id);
                if (!event.crew.includes(crew))
                    event.crew.push({ id, name, role: [role], date: [date]});
                else {
                    crew.role.push(role);
                    crew.date.push(date);
                }
            } else if (eID !== "") {
                events.push({ id: eID, crew: [{ id, name, role: [role], date: [date]}] });
            } else
                return;
        })
    })
    // All unique Events
    let ids = events.map(event => event.id);
    //Retrieving Master Folder
    let year = dateHelper.getYear();
    let yearlyFolder = folder.retrieveFolderByName(`ESS ${year}`);
    let eventsFolder = folder.retrieveFolderByName("1. Events (OTH) - Quotations x Timesheet x Invoice x Payout", yearlyFolder);
    ids.forEach(id => {
        let eventFolder = folder.retrieveFolderByName(id, eventsFolder);
        if (!eventFolder) {
            let hierarchy = {
                parentFolder: id,
                subDirectories: ["Timesheet"]
            };
            folder.generateFolderHierarchy(hierarchy, eventsFolder);
            eventFolder = folder.retrieveFolderByName(id, eventsFolder);
        }
        let timesheetFolder = folder.retrieveFolderByName("Timesheet", eventFolder);
        let f = file.retrieveFileByName(`${id}_Timesheet`, timesheetFolder);
        let fileSS = !f ? file.createSpreadsheetToFolder(`${id}_Timesheet`, timesheetFolder) : SpreadsheetApp.open(f);
        let ss = new Spreadsheet(fileSS);
        if (!f)
            ss.generateTemplate(TemplateType.Timesheet, "Sheet1");
        let modifyData = events.map(event => {return event.id === id ? event.crew : null}).filter(value => value).flat();
        modifyData.sort((a,b) => a.date[a.date.length-1] - b.date[b.date.length-1]);
        ss.insertDataToSheet(TemplateType.Timesheet, "Sheet1", modifyData);
    })
}

/**
 * Generate Individual Template Spreadsheet
 */
function generateIndividualSpreadsheet() {
    let activeSS = SpreadsheetApp.getActive();
    let activeSheet = activeSS.getActiveSheet();
    let sheet = new Sheet(activeSheet);
    let data = sheet.getAllData().slice(1);
    data = data.map(value => value.slice(0,2));

    //Retrieve Individuals Folder
    let rootFolder = folder.retrieveFolderByName("ESS Main Folder");
    let masterFolder = folder.retrieveFolderByName("Individuals",rootFolder);
    let files = folder.retrieveAllSpreadsheetsInFolder(masterFolder);

    let oldMemberFiles = files.filter(file => {
        let fileName = file.getName().split("_");
        let ids = data.map(val => val[0]);
        let names = data.map(val => val[1]);
        return ids.includes(parseInt(fileName[0])) && names.includes(fileName[1])
    });

    let month = dateHelper.getMonthName().slice(0,3);
    let nextMonth = dateHelper.getNextMonth().slice(0,3);
    let {date} = dateHelper.getCurrentDates();
    let year = dateHelper.getYear().slice(2);
    oldMemberFiles.forEach(file => {
        let ss = new Spreadsheet(file);
        if (!ss.getSheet(`${month} ${year}`)) {
            ss.generateTemplate(TemplateType.Individuals,`${month} ${year}`);
        } else {
            if (date > 10) {
                if (!ss.getSheet(`${nextMonth} ${year}`))
                    ss.generateTemplate(TemplateType.Individuals,`${nextMonth} ${year}`);
            }
        }
    });

    let newMember = data.filter(value => {
        let ids = files.map(file => parseInt(file.getName().split("_")[0]));
        let names = files.map(file => file.getName().split("_")[1]);
        if (!value[0] || !value[1])
            return false;
        return !ids.includes(value[0]) && !names.includes(value[1])
    });
    let newMemberFileName = newMember.map(value => value.join("_"));
    newMemberFileName.forEach(member => {
        let f = file.createSpreadsheetToFolder(member,masterFolder);
        let ss = new Spreadsheet(f);
        ss.generateTemplate(TemplateType.Individuals, `${month} ${year}`);
        if (date > 10)
            ss.generateTemplate(TemplateType.Individuals,`${nextMonth} ${year}`);
    })

}