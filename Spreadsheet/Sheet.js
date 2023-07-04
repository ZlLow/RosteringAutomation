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
    }).filter(data => data);
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
      this.insertDataValidation(3, i * 5, dataValidation, 100, 1);

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

    let [month, year] = datestring.split(" ");
    month = month.slice(0,3);
    let days = dateHelper.getDaysInMonth(month);
    let array = Array.from({ length: days }, (_, i) => `${i + 1}`);
    console.log("Inserting Template Header");
    this.insertData(1,1,1,2,[["Date", "Availability"]],true);

    console.log("Inserting Data Validation");
    let dataValidation = ["Available", "till 3pm", "after 3pm","Not Available"];

    array.forEach((value,index) => {
      this.insertData(2 + index,1,1,1,[[`${value} ${month} ${year}`]]);
      this.insertDataValidation(2 + index,2,dataValidation,1,1);
    })

  }

  /**
   * Append Data to the Spreadsheet
   * @param {Array of JSON} externalData: Values retrieved from other Spreadsheet
   * @param {Array} dataValue: Values retrieved from current Sheet
   */
  appendData(externalData, dataValue) {
    console.log("Inserting Modification Data");
    // Find All ID from the current spreadsheet
    let foundData = externalData.filter(dataRow => dataValue.find(valueRow => valueRow[0] === parseInt(dataRow.id))
    );
    if (foundData.length) {
      let headerLength = Object.keys(foundData[0]).length - 1;
      // Remove all None Repeating Data [ID,Name,Mobile,Location];
      let filteredDataValue = dataValue.map(dataRow => dataRow.slice(headerLength));
      filteredDataValue = filteredDataValue.map((row, rowIndex) => { return { row, rowIndex } });

      let modifyData = filteredDataValue.map(data => {
        let rowIndex = data.rowIndex;
        let rowValue = data.row.map((value, colIndex) => {
          if (value === "")
            return colIndex;
        }).filter(value => value !== undefined);
        return { rowValue, rowIndex };
      })
      foundData.sort((a, b) => a.id - b.id);
      modifyData.sort();
      modifyData.forEach(row => {
        let rowIndex = parseInt(row.rowIndex);
        let values = row.rowValue;
        values.forEach(value => {
          let val = parseInt(value);
          let data = foundData[rowIndex].data[val];
          this.insertData(3 + rowIndex, headerLength + (val + 1), 1, 1, [[`${data}`]]);
        })
      });
    }

    console.log("Appending New Data");
    let appendData = externalData.filter(data => !foundData.includes(data)).map(value => [...Object.values(value)].flat());
    appendData.sort();
    if (appendData.length) {
      appendData.forEach(row => this.sheet.appendRow(row))
    }
  }

  /**
   * Update Availability Data by Pulling in data
   * Retrieve Personal Data from Individuals Folder
   * @params {string} folderName: Folder Name to retrieve all the spreadsheet
   * @params {Folder}  rootFolder: Root Folder which folderName resides (Can be empty)
   * @params {string} monthName: Short Name of the Month which is to be inserted into
   */
  insertAvailabilityData(folderName, rootFolder = "", monthName = "") {
    if (typeof folderName !== "string" || (rootFolder && !rootFolder instanceof Object) || (monthName && typeof monthName !== "string"))
      throw new TypeError("Invalid inputs. Please ensure that all the parameters are in the correct format!");
    console.log("Retrieving Personal Data");
    let individualFolder = !rootFolder ?
      folder.retrieveFolderByName(folderName) :
      folder.retrieveFolderByName(folderName, rootFolder);
    let files = folder.retrieveAllSpreadsheetsInFolder(individualFolder);


    console.log("Retrieving All Data from Spreadsheet");
    let personalAvailabilitySS = files.map(file => new Spreadsheet(file));
    if (monthName === "")
      monthName = dateHelper.getMonthName().slice(0, 3);
    let year = dateHelper.getYear().slice(2);
    let availData = personalAvailabilitySS.map(ss => {
      let [id, name, _] = ss.getName().split("_");
      let sheet = ss.getSheet(`${monthName} ${year}`);
      if (sheet) {
        let allData = sheet.getAllData();
        let data = allData.map(data => data[1]).slice(1);
        return { id, name, data };
      } else
        return null;
    }).filter(value => value);
    if (!availData.length) {
      console.log("Unable to retrieve any files");
      return;
    }
    availData.sort();

    console.log("Checking Current Data");
    let ssDataValues = this.getAllData();
    ssDataValues = ssDataValues.slice(2);
    this.appendData(availData, ssDataValues);
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
    }).filter(value => value);
    headerIndex.unshift(0);


    console.log("Retrieving Personal Data");
    let availFile = !rootFolder ?
      file.retrieveFileByName(fileName) :
      file.retrieveFileByName(fileName, rootFolder);
    let availSS = new Spreadsheet(SpreadsheetApp.open(availFile));
    if (monthName === "")
      monthName = dateHelper.getMonthName().slice(0, 3);
    let availSheet = availSS.getSheet(monthName);
    if (!availSheet)
      return;
    let availData = availSheet.getAllData();
    if (!availData.length) {
      console.log("Unable to retrieve any files");
      return;
    }
    var partialAvailability = ["till 3pm", "after 3pm"];
    availData = availData.slice(2);

    console.log("Retrieving from Master Sheet");
    let personalData = masterSheet.getAllData();
    personalData = personalData.slice(1);
    personalData = personalData.filter(dataRow => availData.find(row => row[0] === dataRow[0]) ? true : false);
    personalData = personalData.map(row => row.slice(2, 4))
    availData = availData.map((dataRow, index) => {
      return { id: dataRow[0], name: dataRow[1], mobile: personalData[index][0], location: personalData[index][1], data: dataRow.slice(2) }
    })
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

    console.log("Inserting Data");
    let insertData = args.filter(value => serialNoIndexes.includes(value.id));
    let currentCell = insertRow.filter(row => row[colIndexes[0]]).map((dataRow, index) => {
      let columnIndex = dataRow.slice(nameColIndex[0], dateColIndex + 1).map((col, colIndex) => {
        if (col === "")
          return colIndex;
      }).filter(value => value);
      return { columnIndex, index };
    });
    currentCell.forEach((cellValue, index) => {
      let obj = insertData[index];
      if (obj) {
        let name = insertData[index].name;
        let roles = [... new Set(insertData[index].role)].join(",");
        let dates = insertData[index].date.join(",")
        cellValue.columnIndex.forEach(colValue => {
          console.log(colValue);
          switch (colValue) {
            case 3:
              this.insertData((rowIndex + 2) + index, colValue + colIndexes[0], 1, 1, [[name]]);
              break;
            case 4:
              this.insertData((rowIndex + 2) + index, colValue + colIndexes[0], 1, 1, [[roles]]);
              break;
            case 5:
              this.insertData((rowIndex + 2) + index, colValue + colIndexes[0], 1, 1, [[dates]]);
              break;
            default:
              break;
          }
        })
      }
    })

    let appendData = args.filter(value => !serialNoIndexes.includes(value.id));
    let appendIndex = insertRow.findIndex(row => !row[0]);
    console.log("Appending Data");
    appendData.forEach(value => {
      let dates = value.date.join(",");
      let roles = [...new Set(value.role)].join(",");
      this.insertData(appendIndex + (rowIndex + 1), serialNoColIndex + 1, 1, 1, [[value.id]]);
      this.insertData(appendIndex + (rowIndex + 1), nameColIndex + 1, 1, 1, [[value.name]]);
      this.insertData(appendIndex + (rowIndex + 1), roleColIndex + 1, 1, 1, [[roles]]);
      this.insertData(appendIndex + (rowIndex + 1), dateColIndex + 1, 1, 1, [[dates]]);
      this.insertData(appendIndex + (rowIndex + 1), colIndexes[0] + 1, 1, 1, [[appendIndex++]]);
    })
  }

}