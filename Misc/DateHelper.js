/**
 * Purpose: 1. Used to generate all data related to dates
 *          2. Reduce memories used from creation of new Object
 *          3. Reduce amount of Global Functions
 */
const d = new Date();
const MAX_TIME_INTERVAL = 240000;
const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const dateHelper = class {

    /**
     * Get Current Dates
     * @return {numeric, numeric, numeric} date, month, year : The Current Dates
     */
    static getCurrentDates() {
        const month = d.getMonth();
        const date = d.getDate();
        const year = d.getFullYear();
        return {date, month, year};
    }

    /**
     * Get Date
     * @params {string} datestring: User input date
     * @return {string} date and month
     */
    static getDate(datestring = "") {
        if (typeof datestring !== "string" && !datestring instanceof Date)
            throw new TypeError("There is some errors in converting to string format!");
        let date = !datestring ? d : new Date(datestring);
        let dates = date.getDate();
        let month = months[date.getMonth()].slice(0,3);
        return `${dates} ${month}`;
    }
    /**
     * Get Next Month
     * @return {string} month: Return Name of the month;
     */
    static getNextMonth() {
        let nextMonth = new Date(d.getFullYear(),d.getMonth() + 1,1).getMonth();
        return months[nextMonth];
    }

    /**
     * Get Month Name.
     * @params {string} datestring: User input date
     * @return {string} return name of month.
     * @return {string} return current month name if datestring is empty.
     * @error returns error if wrong type of input
     */
    static getMonthName(datestring = "") {
        if (typeof datestring !== "string" && !datestring instanceof Date)
            throw new TypeError("There is some errors in converting to string format!");
        if (!datestring) {
            console.log("Getting current month");
            return d.toLocaleString('en-US', { month: 'long' });
        }
        let date = Date.parse(datestring);
        return date.toLocaleString('en-US', { month: "long" });
    }
    /**
     * Get Year.
     * @params {string} datestring: User input date
     * @return {string} return year.
     * @return {string} return current year if datestring is empty.
     * @error returns error if wrong type of input
     */
    static getYear(datestring = "") {
        if (typeof datestring !== "string" && !datestring instanceof Date)
            throw new TypeError("There is some errors in converting to string format!");
        if (!datestring) {
            console.log("Getting current year");
            return d.toLocaleString('en-US', { year: 'numeric' });
        }
        let date = Date.parse(datestring);
        return date.toLocaleString('en-US', { year: "numeric" });
    }

    /**
     * Get Number of days in the specific month
     * @params {string} month: User inputted month
     * @return {Array} return number of days in the month
     * @error returns error if wrong type of input
     */
    static getDaysInMonth(month = "") {
        if (typeof month !== "string" && !month instanceof Date)
            throw new TypeError("There is some errors in converting to string format!");
        if (month === "")
            month = this.getMonthName().slice(0,3);
        let shortMonths = months.map(x => x.slice(0,3));
        let monthIndex = shortMonths.findIndex(shortMon => shortMon === month);
        if (monthIndex == -1)
            throw new Error("Month is not in the correct name " + month);
        let year = d.getFullYear();
        return new Date(year,monthIndex+1,0).getDate()
    }

    /**
     * Get Next Year
     * @return {Date} returns the End of Next Year
     */
    static getNextYear() {
        return new Date(this.getYear() + 1, 1, 1);
    }

    static getYearEnd() {
        return new Date(this.getYear(),11,10);
    }
}