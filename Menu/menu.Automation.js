/*------------------------------------------------------------------------------ Menu Trigger ------------------------------------------------------------------------------------------------------*/
/**
 * Trigger Function which updates Roster Spreadsheet
 * Insertion of Data
 */
function updateRosterSpreadsheet() {
    try {
        let activeSS = SpreadsheetApp.getActiveSpreadsheet();
        let activeSheet = activeSS.getActiveSheet();
        let monthName = activeSheet.getName();
        let ss = new Spreadsheet(activeSS);

        ss.insertDataToSheet(TemplateType.Roster, monthName);
    } catch (e) {
        const ui = SpreadsheetApp.getUi();
        const handler = new ErrorHandler(ui);
        handler.createAlert(e);
        ErrorHandler.insertErrorLog(e);
    }
}

/**
 * Trigger Function which refreshes Availability Data
 */
function refreshAvailabilityData(timeBoundEvent) {
    try {
        let startTime = Date.now();
        let activeSS = SpreadsheetApp.getActiveSpreadsheet();
        let activeSheet = timeBoundEvent ? activeSS.getSheetByName(cache.get(`${activeSS.getId()}`)) : SpreadsheetApp.getActiveSheet();
        let triggerName = `${activeSS.getId()}_${activeSheet.getSheetId()}`;
        console.log(activeSheet.getName());
        cache.put(`${activeSS.getId()}`,activeSheet.getName());
        let monthName = activeSheet.getName();
        let year = dateHelper.getYear().slice(2);

        console.log("Retrieving Personal Data");
        let { files } = retrieveIndividualSpreadsheet();

        let keys = properties.getProperty(triggerName);
        console.log(keys);
        if (keys) {
            let trigger = ScriptApp.getProjectTriggers().find(val => val.getUniqueId() === keys);
            if (trigger)
                ScriptApp.deleteTrigger(trigger);
            properties.deleteProperty(triggerName);
        }
        let sheet = new Sheet(activeSheet);
        console.log("Retrieving All Data from Spreadsheet");
        console.log("Running Partitioning");
        files.sort((a,b) => a.getName() - b.getName());
        insertAvailabilityToMain(startTime, triggerName, "refreshAvailabilityData", files,monthName,year,sheet);
        if (Date.now() - startTime > MAX_TIME_INTERVAL && !timeBoundEvent)
            SpreadsheetApp.getUi().alert("Generating User's Template will be running in the background! Please wait!");
    } catch (e) {
        if (!timeBoundEvent) {
            const ui = SpreadsheetApp.getUi();
            const handler = new ErrorHandler(ui);
            handler.createAlert(e);
        } else {
            ErrorHandler.insertErrorLog(e);
        }
    }
}

function clearCache() {
    try {
        let activeSS = SpreadsheetApp.getActiveSpreadsheet();
        let activeSheet = activeSS.getActiveSheet();
        let triggerName = `${activeSS.getId()}_${activeSheet.getSheetId()}`;
        cache.remove(`${triggerName}_Index`);
    } catch (e) {

    }
}

/**
 * Generate Timesheet
 * Servicing of Roster Mastersheet Data into usable Data
 * Get Availability, Roster Role, Event ID, ESS ID & Name
 */
function generateTimesheet(timeBoundEvent) {
    try {
        let startTime = Date.now();
        let activeSS = SpreadsheetApp.getActiveSpreadsheet();
        let activeSheet = timeBoundEvent ? activeSS.getSheetByName(cache.get(`${activeSS.getId()}`)) : SpreadsheetApp.getActiveSheet();
        let triggerName = `${activeSS.getId()}_${activeSheet.getSheetId()}`;
        cache.put(`${activeSS.getId()}`,activeSheet.getName());
        let sheet = new Sheet(activeSheet);
        let data = sheet.getAllData();

        let keys = properties.getProperty(triggerName);
        if (keys) {
            let trigger = ScriptApp.getProjectTriggers().find(val => val.getUniqueId() === keys);
            if (trigger)
                ScriptApp.deleteTrigger(trigger);
            properties.deleteProperty(triggerName);
        }

        let events = cache.get(triggerName);
        if (!events) {
            // Getting ALL Dates
            let dateRange = data[0].filter(value => value !== "");
            dateRange.shift();

            console.log("Getting Subheaders");
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
            console.log("Getting Unique Events");
            events = [];
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
                            event.crew.push({ id, name, role: [role], date: [date] });
                        else {
                            crew.role.push(role);
                            crew.date.push(date);
                        }
                    } else if (eID !== "") {
                        events.push({ id: eID, crew: [{ id, name, role: [role], date: [date] }] });
                    } else
                        return;
                })
            })
        } else
            events = JSON.parse(events);

        //Retrieving Master Folder
        let year = dateHelper.getYear();
        let yearlyFolder = folder.retrieveFolderByName(`ESS ${year}`);
        let eventsFolder = folder.retrieveFolderByName("1. Events (OTH) - Quotations x Timesheet x Invoice x Payout", yearlyFolder);

        console.log("Running Partitioning");
        createTimesheetAndInsert(startTime, triggerName, "generateTimesheet", events, eventsFolder);
        if (Date.now() - startTime > MAX_TIME_INTERVAL && !timeBoundEvent)
            SpreadsheetApp.getUi().alert("Generating User's Template will be running in the background! Please wait!");
    } catch (e) {
        if (!timeBoundEvent) {
            const ui = SpreadsheetApp.getUi();
            const handler = new ErrorHandler(ui);
            handler.createAlert(e);
        } else {
            ErrorHandler.insertErrorLog(e);
        }
    }
}

/**
 * Insert Spreadsheet To all Individual Personnel
 * Work around of Google App Script Execution Time
 * Extend Execution Time of Google App Script from 6 mins to at most 90 mins
 * https://developers.google.com/apps-script/guides/services/quotas
 */
function createNewIndividualTemplate(timeBoundEvent) {
    try {
        let startTime = Date.now();
        let activeSS = SpreadsheetApp.getActiveSpreadsheet();
        let triggerName = `${activeSS.getId()}_create`;
        let activeSheet = SpreadsheetApp.getActiveSheet();
        let sheet = new Sheet(activeSheet);
        let newMemberFile = cache.get(triggerName);
        let { masterFolder, fileNames } = retrieveIndividualSpreadsheet();
        if (!newMemberFile) {
            let data = sheet.getAllData().slice(1);
            data = data.map(value => value.slice(0, 2).join("_"));

            //Retrieve Individuals Folder
            console.log("Retrieve Spreadsheets");
            newMemberFile = data.filter(val => !fileNames.includes(val));
        } else {
            console.log("Getting From Cache");
            newMemberFile = JSON.parse(newMemberFile);
            let keys = properties.getProperty(triggerName);
            if (keys) {
                let trigger = ScriptApp.getProjectTriggers().find(trigger => trigger.getUniqueId() === keys)
                if (trigger)
                    ScriptApp.deleteTrigger(trigger);
                properties.deleteProperty(triggerName);
            }
        }
        insertSpreadsheetToIndividuals(startTime, triggerName, "createNewIndividualTemplate", newMemberFile, masterFolder);
        if (Date.now() - startTime > MAX_TIME_INTERVAL && !timeBoundEvent)
            SpreadsheetApp.getUi().alert("Generating User's Template will be running in the background! Please wait!");
    } catch (e) {
        if (!timeBoundEvent) {
            const ui = SpreadsheetApp.getUi();
            const handler = new ErrorHandler(ui);
            handler.createAlert(e);
        }
        throw e;
    }
}
/**
 * Generate Individual Template Spreadsheet
 */
function generateIndividualSpreadsheet(timeBoundEvent) {
    try {
        let startTime = Date.now();
        let activeSS = SpreadsheetApp.getActiveSpreadsheet();
        let activeSheet = SpreadsheetApp.getActiveSheet();
        let triggerName = `${activeSS.getId()}_generate`;
        let sheet = new Sheet(activeSheet);
        let oldMemberFiles = cache.get(triggerName);
        if (!oldMemberFiles) {
            let data = sheet.getAllData().slice(1);
            data = data.map(value => value.slice(0, 2));
            let dataIDs = data.map(row => row[0]);
            let dataNames = data.map(row => row[1]);

            //Retrieve Individuals Folder
            let { files } = retrieveIndividualSpreadsheet();

            oldMemberFiles = files.filter(file => {
                let fileName = file.getName().split("_");
                return dataIDs.includes(parseInt(fileName[0])) && dataNames.includes(fileName[1])
            });
            oldMemberFiles = oldMemberFiles.map(file => file.getId());
        } else {
            oldMemberFiles = JSON.parse(oldMemberFiles);
            let keys = properties.getProperty(triggerName);
            if (keys) {
                let trigger = ScriptApp.getProjectTriggers().find(trigger => trigger.getUniqueId() === keys)
                if (trigger)
                    ScriptApp.deleteTrigger(trigger);
                properties.deleteProperty(triggerName);
            }
        }
        insertSpreadsheetToIndividuals(startTime, triggerName, "generateIndividualSpreadsheet", oldMemberFiles);
        if (Date.now() - startTime > MAX_TIME_INTERVAL && !timeBoundEvent)
            SpreadsheetApp.getUi().alert("Generating User's Template will be running in the background! Please wait!");
    } catch (e) {
        if (!timeBoundEvent) {
            const ui = SpreadsheetApp.getUi();
            const handler = new ErrorHandler(ui);
            handler.createAlert(e);
        }
        throw e;
    }
}

/**
 * Retrieve All Individuals Spreadsheets
 */
function retrieveIndividualSpreadsheet() {
    let rootFolder = folder.retrieveFolderByName("ESS Main Folder");
    let masterFolder = folder.retrieveFolderByName("Individuals", rootFolder);
    let files = folder.retrieveAllSpreadsheetsInFolder(masterFolder);
    let fileNames = files.map(file => file.getName());

    return { masterFolder, files, fileNames };
}

/*--------------------------------------------------------------------------Time Based Execution ---------------------------------------------------------------------------------------------------- */

/**
 * @params {number} startTime: Starting Time of the first Execution;
 * @params {String} triggerName: Name of the Trigger For Easier Trigger Management
 * @params {String} functionName: Name of the function to be triggered for;
 * @params {Array<String>} fileArrays: Array of string that contains either File Name (Create) or File ID (Generate)
 * @params {Folder} rootFolder: Root Folder in which the spreadsheet is to be inserted
 */
function insertSpreadsheetToIndividuals(startTime, triggerName, functionName, fileArrays, rootFolder = "") {
    // Retrieve Month & Year
    let month = dateHelper.getMonthName().slice(0, 3);
    let nextMonth = dateHelper.getNextMonth().slice(0, 3);
    let { date } = dateHelper.getCurrentDates();
    let year = dateHelper.getYear().slice(2);

    console.log("Partitioning All Spreadsheets");

    // Read Access
    for (let i = 0; i < fileArrays.length; i++) {
        if (Date.now() - startTime < MAX_TIME_INTERVAL) {
            let fileSS = fileArrays[i];
            let f = !rootFolder ? SpreadsheetApp.openById(fileSS) : file.createSpreadsheetToFolder(fileSS, rootFolder);
            let ss = new Spreadsheet(f);
            if (!ss.getSheet(`${month} ${year}`))
                ss.generateTemplate(TemplateType.Individuals, `${month} ${year}`);
            if (date > 10)
                ss.generateTemplate(TemplateType.Individuals, `${nextMonth} ${year}`);
        } else {
            console.log("Extending Time Through Triggers");
            cache.put(triggerName, JSON.stringify(fileArrays.slice(i)), 3600);
            timeBasedTrigger(triggerName, functionName);
            return;
        }
    }
    cache.put(triggerName, JSON.stringify([]));
}



/**
 * Insert Spreadsheet To all Individual Personnel
 * Work around of Google App Script Execution Time
 * Extend Execution Time of Google App Script from 6 mins to at most 90mins
 * https://developers.google.com/apps-script/guides/services/quotas
 *
 * @params {number} startTime: Starting Time of the first Execution
 * @params {String} triggerName: Name of the Trigger For Easier Trigger Management
 * @params {String} functionName: Name of the function to be triggered for
 * @params {JSON} events: JSON representation of Events
 * @params {Folder} rootFolder: Root Folder in which the spreadsheet is to be inserted
 */
function createTimesheetAndInsert(startTime, triggerName, functionName, events, rootFolder = "") {
    for (let i = 0; i < events.length; i++) {
        if (Date.now() - startTime < MAX_TIME_INTERVAL) {
            let event = events[i];
            let id = event.id;
            let eventFolder = folder.retrieveFolderByName(id, rootFolder);
            if (!eventFolder) {
                let hierarchy = {
                    parentFolder: id,
                    subDirectories: ["Timesheet"]
                };
                folder.generateFolderHierarchy(hierarchy, rootFolder);
                eventFolder = folder.retrieveFolderByName(id, rootFolder);
            }
            let timesheetFolder = folder.retrieveFolderByName("Timesheet", rootFolder);
            let f = file.retrieveFileByName(`${id}_Timesheet`, timesheetFolder);
            let fileSS = !f ? file.createSpreadsheetToFolder(`${id}_Timesheet`, timesheetFolder) : SpreadsheetApp.open(f);
            let ss = new Spreadsheet(fileSS);
            if (!f)
                ss.generateTemplate(TemplateType.Timesheet, "Sheet1");
            let modifyData = event.crew;
            if (!modifyData)
                continue;
            modifyData.sort((a, b) => a.date.length - b.date.length);
            ss.insertDataToSheet(TemplateType.Timesheet, "Sheet1", modifyData);
        } else {
            cache.put(triggerName, JSON.stringify(events.slice(i)), 3600);
            timeBasedTrigger(triggerName, functionName);
            return;
        }
    }
}

/**
 *
 * @params {number} startTime: Starting Time of the first Execution
 * @params {String} triggerName: Name of the Trigger For Easier Trigger Management
 * @params {String} functionName: Name of the function to be triggered for
 * @params {Array} fileArrays: Array of files
 * @params {Sheet} masterSheet: Availability Sheet
 * @params {string} month: Short Name of the Month
 * @params {number} year: Year which is being retrieved
 */
function insertAvailabilityToMain(startTime, triggerName, functionName, fileArrays, month, year, masterSheet) {
    console.log("Partitioning All Spreadsheets");
    let files = [];
    let queueIndex = cache.get(`${triggerName}_Index`);
    console.log(queueIndex);
    queueIndex = !queueIndex ? 0 : JSON.parse(queueIndex);
    for (queueIndex; queueIndex < fileArrays.length; queueIndex++) {
        if (Date.now() - startTime < MAX_TIME_INTERVAL) {
            let fileID = fileArrays[queueIndex].getId();
            let [id, name] = fileArrays[queueIndex].getName().split("_");
            console.log(id,name);
            let f = SpreadsheetApp.openById(fileID);
            let ss = new Spreadsheet(f);
            let sheet = ss.getSheet(`${month} ${year}`);
            if (!sheet)
                continue;
            let data = sheet.getAllData().slice(1,32).flatMap(val => val.slice(1));
            files.push({id, name, data})
        } else {
            console.log("Extending Time Through Triggers");
            cache.put(`${triggerName}_Index`, queueIndex, 3600);
            timeBasedTrigger(triggerName, functionName);
            break;
        }
    }
    masterSheet.insertAvailabilityData(files);
    return;
}