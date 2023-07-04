const file = class {
    /**
     * Retrieve a specific File By Name
     * Does not check the based on the parent folder
     * @params {string} fileName : Name of the specific file
     * @params {Folder} rootFolder : The Parent Folder which the specific file exist in
     * @returns {file} returns first file or null if unable to retrieve folder
     */
    static retrieveFileByName(fileName, rootFolder = "") {
        if (typeof fileName !== "string" || (rootFolder && !rootFolder instanceof Object))
            throw new TypeError("There is some errors in retrieving the file!");
        if (!rootFolder) {
            var fileIterator = DriveApp.getFilesByName(fileName)
            var file = fileIterator.hasNext() ? fileIterator.next() : null;
        }
        else
            file = miscTools.recursiveSearchFile(fileName,rootFolder);
        return file;
    }
    /**
     * Creates File to Folder
     * @params {string} fileName: Name of the file
     * @params {Folder} rootFolder: The Parent Folder which the specific file is being created
     * @returns {file} returns the file that have been created
     */
    static createSpreadsheetToFolder(fileName, rootFolder = "") {
        if (typeof fileName !== "string" || (rootFolder && !rootFolder instanceof Object))
            throw new TypeError("There is some errors in retrieving the file!");
        let ss = SpreadsheetApp.create(fileName);
        !rootFolder ? DriveApp.getFileById(ss.getId()) : DriveApp.getFileById(ss.getId()).moveTo(rootFolder);
        return ss;
    }

}