class Sheet {
  constructor(sheet) {
    this.sheet = sheet
  }

  /**
   * Get All Data from the specific Sheet
   *
   * @returns {2D Array} Returns all values
   * Example: [
   *           ["Refresh Data","Generate Data"],
   *           ["no","no","ESS Roster Sheet","",...].
   *           ["ESS ID","Name"],
   *           ...
   *          ]
   */
  getAllData() {
    let range = this.sheet.getDataRange();
    return range.getValues();
  }

  /**
   * Get the Header of the Spreadsheet
   * @params {string} headerValue: Retrieve Header Value
   * @return {index, colIndex} returns the row and column of the index
   * Date Sample:
   * [
   * {rowIndex: 0, colIndexes: [1,2,3,4]}
   * ]
   */
  getHeader(headerValue) {
    let allData = this.getAllData();
    return allData.map((row, rowIndex) => {
      let colIndexes = row.map((value, colIndex) => {
        if (value === headerValue)
          return colIndex;
      })
          .filter(value => value !== undefined);
      if (colIndexes.length)
        return { rowIndex, colIndexes };
    }).filter(data => data !== undefined);
  }

  /**
   * Insert Data into Sheet
   * @params {numbers} startingRowIndex: Row Index
   * @params {numbers} startingColIndex: Column Index
   * @params {numbers} numsRow: Number of Rows
   * @params {numbers} numsCol: Number of Columns
   * @params {2D Array} data: Values of the specific cell
   * data Example: [
   *                ["Header1","Header2"],
   *                ["Data1","Data2"]
   *               ]
   * @params {Array} ...args: Represent any leftover parameters
   * ...args Current Parameters:
   *         [0]: {Boolean} setBold: set bold
   *         [1]: {Boolean} setMerge: set merged range
   *         [2]: {Boolean} setCheckbox :set specific cell to checkbox
   *         [3]: {Boolean} setCenterAlignment:  center alignment cell
   *         [4]: {Boolean} setWrap: set text wrap
   * @return {Range} returns the specific range from Starting Row Index & Column Index
   */
  insertData(startingRowIndex, startingColIndex, numsRow, numsCol, data, ...args) {
    if (typeof startingRowIndex !== "number" || typeof startingColIndex !== "number")
      throw new TypeError("An incorrect type have been entered! : Ensure that startingRowIndex & startingColIndex is numeric");
    var range = this.sheet.getRange(startingRowIndex, startingColIndex, numsRow, numsCol);
    if (data.length)
      range.setValues(data);
    if (args.length) {
      var setBold = args[0]
      var setMerge = args[1];
      var setCheckbox = args[2];
      var setCentered = args[3];
      var setWrap = args[4];
      if (setBold !== undefined && setBold)
        range.setFontWeight("bold");
      if (setMerge !== undefined && setMerge)
        range.merge();
      if (setCheckbox !== undefined && setCheckbox)
        range.insertCheckboxes("yes", "no");
      if (setCentered !== undefined && setCentered)
        range.setHorizontalAlignment("center");
      if (setWrap !== undefined && setWrap)
        range.setWrap(true);
    }
    return range;
  }

  /**
   * Insertion of Data Validation
   * @params {numbers} startingRowIndex: Row Index
   * @params {numbers} startingColIndex: Column Index
   * @params {numbers} numsRow: Number of Rows
   * @params {numbers} numsCol: Number of Columns
   * @params {Array} validationData : List of Values for Data Validation
   * Example of validationData: ["Data1","Data2","Data3"]
   * @returns {Range} returns range
   */
  insertDataValidation(startingRowIndex, startingColIndex, validationData, numsRow, numsCol) {
    if (typeof startingRowIndex !== "number" || typeof startingColIndex !== "number" && (!Array.isArray(validationData)))
      throw new TypeError("An incorrect type have been entered! : Ensure that startingRowIndex & startingColIndex is numeric and validationData is an array");
    var rule =
        SpreadsheetApp.newDataValidation()
            .requireValueInList(validationData)
            .build();
    var range = this.sheet.getRange(startingRowIndex, startingColIndex, numsRow, numsCol);
    return range.setDataValidation(rule);
  }

  /**
   * Generate Availability Template
   * Inserts Checkboxes, Headers and Dates
   * Does not insert Personal Data
   * @params {string} monthName: Name of the month
   */
  createAvailabilityTemplate(monthName = "") {
    if (typeof monthName !== "string")
      throw new TypeError("Ensure that name of the month is in string format");
    console.log("Inserting Header");
    this.insertData(2, 1, 1, 2, [["ESS ID", "Name"]], true);

    console.log("Inserting Dates");
    if (monthName === "") {
      monthName = dateHelper.getMonthName();
      var days = dateHelper.getDaysInMonth();
    } else
      days = dateHelper.getDaysInMonth(monthName);
    let dateData = [Array.from({ length: days }, (_, i) => `${i + 1} ${monthName}`)];
    let emptyCol = Array.from({ length: days - 1 }, x => "");
    this.insertData(2, 3, 1, days, dateData, true);
    this.insertData(1, 3, 1, days, [["Date"].concat(emptyCol)], true, true, false);
  }

  /**
   * Generate Roster Template
   * Inserts Checkboxes, Headers, Dates, SubHeader of Dates and Data Validation
   * Does not insert Personal Data
   * @params {string} monthName: Name of the month
   */
  createRosterTemplate(monthName = "") {
    if (typeof monthName !== "string")
      throw new TypeError("Ensure that name of the month is in string format");
    console.log("Inserting Headers");
    this.insertData(2, 1, 1, 4, [["ESS ID", "Name", "Mobile", "Indicate the area that you are living in. (e.g. Jurong, Sengkang, Woodlands, Tampines, etc)"]], true, false, false, true, true);
    this.insertData(1, 1, 1, 4, [["ESS Roster Sheet", "", "", ""]], true, true);

    console.log("Inserting Dates with Sub Headers");
    if (monthName === "") {
      monthName = dateHelper.getMonthName();
      var days = dateHelper.getDaysInMonth();
    } else {
      days = dateHelper.getDaysInMonth(monthName);
    }
    var dateHeaderData = [...Array(days)];
    dateHeaderData.fill([
      "Availability",
      "Partially Available (e.g. Free Till 3pm OR Free AFTER 3pm.)",
      "Rostered Role",
      "Rostered For (Event Name)",
      "Event ID"
    ]);
    this.insertData(2, 5, 1, days * 5, [dateHeaderData.flat()], true, false, false, true, true);

    console.log("Inserting Data Validation Rules");
    let dataValidation = ["Available", "Not Available", "Others"];
    for (let i = 1; i <= dateHeaderData.length; i++) {
      let borderedRange = this.sheet.getRange(1, i * 5, 999, 1);
      borderedRange.setBorder(null, true, null, null, null, null);
      this.insertDataValidation(3, i * 5, dataValidation, 500, 1);

      this.insertData(1, i * 5, 1, 5, [[`${i} ${monthName}`, "", "", "", ""]], true, true, false, true, true);
    }
  }

  /**
   * Generate Timesheet Template
   * Inserts Headers, Dates
   */
  createTimesheetTemplate() {
    let placeholder = Array.from({ length: 99 }, (x) => [""]);
    console.log("Inserting Headers");
    this.insertData(1, 1, 1, 5, [
      ["Input Event Code (Refer to Master Account Sheet)", "", "", "", ""]
    ], true, true, false, true, true)
        .setBorder(true, true, true, true, null, null)
        .setBackground("#93ccea");

    this.insertData(2, 1, 1, 5, [["Valuation Data (Refer to Bank Statement)", "", "", "", ""]], true, true, false, true, true)
        .setBorder(true, true, true, true, null, null)
        .setBackground("#a4c2f4");

    this.insertData(3, 1, 1, 5, [["Done By (Input your Name)", "", "", "", ""]], true, true, false, true, true)
        .setBorder(true, true, true, true, null, null)
        .setBackground("#a4c2f4");

    let approval = ["P. Approval", "Approval"];
    let invoice = ["Invoice Not Sent", "Invoice Sent"];

    this.insertDataValidation(1, 6, approval, 1, 1);
    this.insertDataValidation(1, 7, invoice, 1, 1);

    let rightHeader = this.insertData(2, 6, 2, 1, [[""], ["(Input your Name)"]], true, false, false, true, true);
    let subHeader = this.insertData(4, 1, 1, 16, [["", "No.", "ESS Serial No.", "Name", "Role", "Date", "Start Time", "End Time", "Duration", "Rate", "Amount", "Transport Claims", "Meals Claims", "ART Claims", "Amount", "Remarks"]], true, false, false, true, true);

    console.log("Inserting Borders & Colors");
    rightHeader.setBorder(true, true, true, true, null, null);
    rightHeader.setBackground("#93c47d");

    subHeader.setBorder(true, true, true, true, true, true);
    // HexaColor for Light Cornflower Blue
    subHeader.setBackground("	#a4c2f4");

    console.log("Inserting Columns for Users")
    let duration = this.insertData(5, 9, 100, 1, [[""], ...placeholder], false, false, false, true, true);
    duration.setFormulaR1C1("=(INT(R[0]C[-1]/100)+(MOD(R[0]C[-1],100)/60))-(INT(R[0]C[-2]/100)+(MOD(R[0]C[-2],100)/60))");

    console.log("Inserting Data Validation");
    let hourlyRate = [12, 14, 18, 20];
    this.insertDataValidation(5, 10, hourlyRate, 100, 1);

    let rateTotal = this.insertData(5, 11, 100, 1, [[""], ...placeholder], false, false, false, true, true);
    rateTotal.setFormulaR1C1("=R[0]C[-1]*R[0]C[-2]")
        .setNumberFormat("$0.00");

    let totalAmounts = this.insertData(5, 15, 100, 1, [[""], ...placeholder]);
    totalAmounts.setFormulaR1C1("=R[0]C[-1]+R[0]C[-2]+R[0]C[-3]+R[0]C[-4]")
        .setNumberFormat("$0.00");

    console.log("Inserting Column Header");
    this.insertData(5, 1, 100, 1, [["Enter Event Title (Refer to Master Account Sheet)"], ...placeholder], true, true, false, true, true);

    console.log("Inserting Footer");
    let footer = this.insertData(105, 1, 1, 16, [["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]]);
    footer.setBackground("#a4c2f4");

    console.log("Inserting Total Amounts");
    this.insertData(106, 10, 1, 1, [["Total"]], true)
        .setBackground("#a4c2f4")
        .setBorder(true, true, true, true, null, null);
    let totalSum = this.insertData(106, 11, 1, 1, [[""]], true);
    totalSum.setFormulaR1C1("=SUM(R[-102]C[0]:R[-2]C[0])")
        .setBorder(true, true, true, true, null, null)
        .setNumberFormat("$0.00");

    console.log("Inserting Total with Claims");
    this.insertData(106, 14, 1, 1, [["Total Claims"]], true)
        .setBackground("#a4c2f4")
        .setBorder(true, true, true, true, null, null);
    let totalClaims = this.insertData(106, 15, 1, 1, [[""]], true);
    totalClaims.setFormulaR1C1("=SUM(R[-102]C[0]:R[-2]C[0])")
        .setBorder(true, true, true, true, null, null)
        .setNumberFormat("$0.00");

    console.log("Inserting Row Example");
    this.insertData(108, 2, 1, 15, [["", "", "", "Copy Paste this whole row for every New Entries/Rows or inserts last row", "", "", "", "", "", "", "", "", "", "", 0]], true, false, false, false, true)
        .setBorder(true, true, true, true, null, null);

    console.log("Inserting Formulas for Duration");
    let sampleDurations = this.insertData(108, 9, 1, 1, [[""]], false, false, false, true, true);
    sampleDurations.setFormulaR1C1("=(INT(R[0]C[-1]/100)+(MOD(R[0]C[-1],100)/60))-(INT(R[0]C[-2]/100)+(MOD(R[0]C[-2],100)/60))");

    console.log("Inserting Formulas for Total Rate Amount");
    let samplerateTotal = this.insertData(108, 11, 1, 1, [[""]], false, false, false, true, true);
    samplerateTotal.setFormulaR1C1("=R[0]C[-1]*R[0]C[-2]")
        .setNumberFormat("$0.00");

    console.log("Inserting Formulas for Total Amount");
    let sampletotalAmounts = this.insertData(108, 15, 1, 1, [[""]]);
    sampletotalAmounts.setFormulaR1C1("=R[0]C[-1]+R[0]C[-2]+R[0]C[-3]+R[0]C[-4]")
        .setNumberFormat("$0.00");


    console.log("Inserting Border for all Rows");
    var dataRange = this.sheet.getRange(4, 1, 102, 16);
    dataRange.setBorder(true, true, true, true, true, true);
  }

  /**
   * Create Template for Individual Spreadsheet
   * @param {string} datestring: Name of the sheet
   */
  createIndividualTemplate(datestring = "") {
    if (typeof datestring !== "string")
      throw new TypeError("Ensure that name of the month is in string format");
    let month, year;
    if (datestring) {
      [month, year] = datestring.split(" ");
      month = month.slice(0, 3);
    } else {
      month = dateHelper.getMonthName().slice(0, 3);
      year = dateHelper.getYear();
    }
    let days = dateHelper.getDaysInMonth(month);
    let array = Array.from({ length: days }, (_, i) => `${i + 1} ${month} ${year}`);
    console.log("Inserting Template Header");
    this.insertData(1, 1, 1, 2, [["Date", "Availability"]], true);

    let dataArray = array.map(value => [value]);
    console.log("Inserting Data Validation");
    let dataValidation = ["Available", "till 3pm", "after 3pm", "Not Available"];
    this.insertData(2, 1, array.length, 1, dataArray);
    this.insertDataValidation(2, 2, dataValidation, array.length, 1);
    SpreadsheetApp.flush();
  }

  /**
   * Append Data to the Spreadsheet
   * @param {Array of JSON} externalData: Values retrieved from other Spreadsheet
   * @param {Array} dataValue: Values retrieved from current Sheet
   * @param {boolean} canReplace: Value that can replace other value;
   */
  appendData(externalData, dataValue, canReplace = false) {
    // Find All ID from the current spreadsheet
    let dataIDs = dataValue.map(val => val[0]);
    let foundData = externalData.filter(dataRow => dataIDs.includes(parseInt(dataRow.id)));
    if (foundData.length) {
      console.log("Inserting Modification Data");
      let headerLength = 2;

      let modifyData = foundData.map(val => [...Object.values(val)].flat().slice(headerLength));
      let currentData = dataValue.map(data => data.slice(headerLength));
      if (!canReplace) {
        currentData.forEach((row,rowIndex) => {
          modifyData[rowIndex] = row.map((val, col) => (val === "" && modifyData[rowIndex][col] ? modifyData[rowIndex][col] : val));
        })
      }
      this.insertData(3, headerLength + 1, modifyData.length, modifyData[0].length, modifyData);

    }
    let lastRow = this.sheet.getLastRow();
    console.log("Appending New Data");
    let appendData = externalData.filter(data => !foundData.includes(data)).map(value => [...Object.values(value)].flat());
    if (appendData.length)
      this.insertData(lastRow + 1, 1, appendData.length, appendData[0].length, appendData);
    return;
  }

  /**
   * Update Availability Data by Pulling in data
   * @params {string} sheetName: Name of the Sheet
   * @params {Array} args: Data in which Availability Spreadsheet Require to Service
   * [
   *  {id: 1 , name: "David" ,data: ['Available','Not Available']},
   *  {id: 2 , name: "Yi Xin" ,data: ['Not Available','Not Available']},
   *  {id: 3 , name: "Aaron" ,data: ['Available','Available']}
   * ]
   */
  insertAvailabilityData(args=[]) {
    if ((args && !args instanceof Array))
      throw new TypeError("Invalid inputs. Please ensure that all the parameters are in the correct format!");
    let availData = args;
    availData.sort((a, b) => parseInt(a.id) - parseInt(b.id));
    console.log("Checking Current Data");
    let ssDataValues = this.getAllData();
    ssDataValues = ssDataValues.slice(2);
    this.appendData(availData, ssDataValues, true);
    return;
  }

  /**
   * Update Roster Data by Pulling in data
   * Retrieve Availability Data
   * @params {string} fileName: Specific File Name that retrieve all Availability Data
   * @params {Sheet} masterSheet: The master sheet which contains all info about Personnel
   * @params {Folder}  rootFolder: Root Folder which folderName resides (Can be empty)
   * @params {string} monthName: Short Name of the Month which is to be inserted into
   */
  insertRosterData(fileName, masterSheet, rootFolder = "", monthName = "") {
    if (typeof fileName !== "string" || (rootFolder && !rootFolder instanceof Object) || (monthName && typeof monthName !== "string"))
      throw new TypeError("Invalid inputs. Please ensure that all the parameters are in the correct format!");
    console.log("Checking Current Data");
    let ssDataValues = this.getAllData();
    let headerRow = ssDataValues[0];
    ssDataValues = ssDataValues.slice(2);
    let headerIndex = headerRow.map((value, index) => {
      if (value !== "")
        return index
    }).filter(value => value !== undefined);

    console.log("Retrieving Personal Data");
    let availFile = !rootFolder ?
        file.retrieveFileByName(fileName) :
        file.retrieveFileByName(fileName, rootFolder);
    let availSS = new Spreadsheet(SpreadsheetApp.open(availFile));
    if (monthName === "")
      monthName = dateHelper.getMonthName().slice(0, 3);
    let availSheet = availSS.getSheet(monthName);
    if (!availSheet)
      throw new Error("Unable to retrieve the sheet");
    let availData = availSheet.getAllData();
    if (!availData.length) {
      console.log("Unable to retrieve any files");
      throw new Error("Empty Sheet is found");
    }
    var partialAvailability = ["till 3pm", "after 3pm"];
    availData = availData.slice(2);
    let dataIDs = availData.map(row => row[0]);

    console.log("Retrieving from Master Sheet");
    let personalData = masterSheet.getAllData();
    personalData.shift();
    personalData = personalData.filter(dataRow => dataIDs.includes(dataRow[0]));
    if (personalData.length) {
      personalData = personalData.map(row => row.slice(2, 4));
      availData = availData.map((dataRow, index) => {
        return { id: dataRow[0], name: dataRow[1], mobile: personalData[index][0], location: personalData[index][1], data: dataRow.slice(2) }
      }).filter(val => val !== undefined);
    } else {
      availData = availData.map(dataRow => {
        return { id: dataRow[0], name: dataRow[1], data: dataRow.slice(2)}
      });
    }
    availData.sort((a, b) => a.id - b.id);
    availData.forEach(dataRow => {
      headerIndex.forEach((header, i) => {
        let diff = i === headerIndex.length - 1 ? 0 : i === 0 ? headerIndex[1] - headerIndex[0] : (headerIndex[i + 1] - header) - 1;
        let emptyCol = [...Array(diff).keys()].fill("");
        let startIndex = i === 0 ? header + 1 : header + 2;
        dataRow.data.splice(startIndex, 0, ...emptyCol);
        let avail = dataRow.data[startIndex - 1];
        if (partialAvailability.includes(avail)) {
          dataRow.data[startIndex - 1] = "Others";
          dataRow.data[startIndex] = avail;
        }
      })
    })
    this.appendData(availData, ssDataValues);
  }
  /**
   * Inserting Data into Timesheet
   * @params {Array of JSON} args: Arguments for creation of Timesheet Data
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
   * Different Dates can have repeating Events
   */
  insertTimesheetData(args) {
    let { rowIndex, colIndexes } = this.getHeader("No.")[0];
    let serialNoColIndex = this.getHeader("ESS Serial No.")[0].colIndexes[0];
    let nameColIndex = this.getHeader("Name")[0].colIndexes[0];
    let roleColIndex = this.getHeader("Role")[0].colIndexes[0];
    let dateColIndex = this.getHeader("Date")[0].colIndexes[0];
    let insertRow = this.getAllData().slice(rowIndex + 1);
    let serialNoIndexes = insertRow.map(data => data[serialNoColIndex]);
    insertRow = insertRow.map(val => val.slice(1));

    console.log("Inserting Data");
    let data = args.filter(value => serialNoIndexes.includes(value.id));
    if (data.length) {
      let currentCell = insertRow.filter(row => row[colIndexes[0]]).map((dataRow, rowIndex) => {
        let columnIndexes = dataRow.slice(nameColIndex[0], dateColIndex).map((col, colIndex) => {
          if (col === "")
            return colIndex;
        }).filter(value => value !== undefined);
        return { columnIndexes, rowIndex };
      }).filter(val => val.columnIndexes.length);
      currentCell.forEach(row => {
        let rowIndex = row.rowIndex;
        let val = data[rowIndex];
        if (val) {
          let roles = [... new Set(val.role)].join(",");
          let dates = [... new Set(val.date)].join(",");
          row.columnIndexes.forEach(col => {
            if (col == serialNoColIndex - 1) insertRow[rowIndex][col] = val.id;
            else if (col == nameColIndex - 1) insertRow[rowIndex][col] = val.name;
            else if (col == roleColIndex - 1) insertRow[rowIndex][col] = roles;
            else if (col == dateColIndex - 1) insertRow[rowIndex][col] = dates;
            return;
          })
        }
      })
      this.insertData(rowIndex + 2,colIndexes[0] + 1, insertRow.length,insertRow[0].length, insertRow);
    }

    console.log("Appending Data");
    let appendIndex = insertRow.findIndex(row => !row[0]);
    let appendData = args.filter(value => !serialNoIndexes.includes(value.id))
        .map((val,index) => {
          let roles = [... new Set(val.role)].join(",");
          let dates = [... new Set(val.date)].join(",");
          return [appendIndex + index+ 1,val.id,val.name,roles,dates];
        });
    if (appendData.length)
      this.insertData(appendIndex + (rowIndex + 2), colIndexes[0] + 1,appendData.length,appendData[0].length, appendData);
  }
}