class Menu {

    /**
     * Template Create Menu
     * @params {Spreadsheet} spreadsheet: Spreadsheet which can retrieve menu from
     * @params {string} menuName: Name of the main Menu
     * @params {Array} args: Menu items which contains both name of the menu and name of the function
     * Example args: [["subMenuName","functionName"],[...args]]
     */
    static createMenu(spreadsheet, menuName, ...args) {
        if (spreadsheet instanceof Object &&  typeof menuName !== "string" && !Array.isArray(args))
            throw new TypeError("Inputs are invalid, please ensure that it follow the correct format");
        let data = args.map(row => {return {name: row[0], functionName: row[1]}})
        console.log(data);
        spreadsheet.addMenu(menuName, data);
    }
}