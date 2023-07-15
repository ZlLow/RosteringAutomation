class ErrorHandler {
    constructor(ui) {
        this.ui = ui;
    }

    /**
     * Insert Error Log into Spreadsheet
     * @params {Error} error: Error received
     */
    static insertErrorLog(error) {
        let year = dateHelper.getYear();
        let masterFolder = folder.retrieveFolderByName("ESS Main Folder");
        let yearFolder = folder.retrieveFolderByName(`ESS ${year}`,masterFolder);
        let fileSS = file.retrieveFileByName("Error Log",yearFolder) ? SpreadsheetApp.open(file.retrieveFileByName("Error Log",yearFolder)) : file.createSpreadsheetToFolder("Error Log",yearFolder);
        let ss = new Spreadsheet(fileSS);
        let sheet = ss.getSheet("Sheet1");
        sheet.sheet.appendRow([error.stack,error]);
    }

    /**
     * Create Alert Using Spreadsheet UI
     * @params {Error} error: Error received
     */
    createAlert(error) {
        const button = this.ui.ButtonSet.OK_CANCEL;
        console.log(error);
        switch (error) {
            case SyntaxError:
            case ReferenceError:
            case TypeError:
                console.log(error.stack);
                this.ui.alert("Error","Internal Error. Please wait for awhile and try again later!",button);
                break;
            case RangeError:
                console.log(error.stack);
                this.ui.alert("Error","Please ensure that the HEADER NAMING CONVENTION HAS NOT CHANGED",button);
                break;
            default:
                console.log(error.stack);
                this.ui.alert("Error", error,button);
        }
    }
}