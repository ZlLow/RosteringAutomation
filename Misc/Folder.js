const folder = class {
    /**
     * Generate Whole Folder Hierarchy
     * Only Creates Subfolders within the root directory
     * @params {Array of Object} subFolders: An array that contains the parent folder name and a list of subdirectories' name
     * subFolder Example: {parentFolder: rootFolderName, subDirectories: [folder1,folder2,folder3], args: {parentFolder : folder1, subDirectories: [folder4], args: {...args}}}
     * @params {Folder} rootFolder: Parent Folder of the Folder Hierarchy
     */
    static generateFolderHierarchy(subFolder, rootFolder = "") {
        if ((!subFolder instanceof Object || (rootFolder && !rootFolder instanceof Object)))
            throw new TypeError("Invalid input. Ensure that subFolder follows the format!");
        console.log("Retrieving From Master Folder");
        let parent = !rootFolder ? this.retrieveFolderByName(subFolder.parentFolder) : this.retrieveFolderByName(subFolder.parentFolder, rootFolder);
        console.log("Creating Parent Folder");
        if (!parent)
            parent = !rootFolder ? DriveApp.createFolder(subFolder.parentFolder) : rootFolder.createFolder(subFolder.parentFolder);
        var folderArray = [];
        console.log("Creating Sub Directories")
        subFolder.subDirectories.forEach(name => {
            let child = this.retrieveFolderByName(name, parent);
            !child ?
                folderArray.push(parent.createFolder(name)) :
                folderArray.push(child);
        });
        let args = subFolder.args;
        if (!args)
            return;
        console.log("Creating Directories within Sub Directories")
        miscTools.recursiveFolderCreation(args, folderArray);
    }

    /**
     * Retrieve a specific Folder By Name
     * Does not check the based on the parent folder
     * @params {string} folderName : Name of the specific folder
     * @params {Folder} rootFolder : The Parent Folder which the specific folder exist in
     * @returns {Folder} returns first folder or null if unable to retrieve folder
     */
    static retrieveFolderByName(folderName, rootFolder = "") {
        if (typeof folderName !== "string" || (rootFolder && !rootFolder instanceof Object))
            throw new TypeError("There is some errors in retrieving the folder!");
        if (!rootFolder) {
            var folderIterator = DriveApp.getFoldersByName(folderName)
            var folder = folderIterator.hasNext() ? folderIterator.next() : null;
        }
        else
            folder = miscTools.recursiveSearchFolder(folderName, rootFolder);
        return folder;
    }

    /**
     * Retrieve all files
     * @params {folder} rootFolder: The folder that contains all the spreadsheets
     * @return {Array[File]} return an array of files within the folder
     */
    static retrieveAllSpreadsheetsInFolder(rootFolder) {
        if (!rootFolder instanceof Object)
            throw new TypeError("There is some errors in the parameter: rootFolder is not in the right format!");
        let files = [];
        let iterator = rootFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
        while (iterator.hasNext()) {
            files.push(iterator.next());
        }
        return files;
    }

    /**
     * Retrieve Folder Hierarchy of specific folder
     * @params {Folder} rootFolder: Targeted Folder
     * @return {Objects} return as an JSON formatted Object
     * Return Value Example: {id: rootFolderID, name:rootFolderName, subDirectories: {id: folder1ID, name: folder1Name, subDirectories: {...args}}}
     */
    static retrieveFolderHierarchy(rootFolder) {
        if (!rootFolder instanceof Object)
            throw new TypeError("There is some errors in the parameters: rootFolder is not in the right format!");
        return miscTools.recursiveFolder(rootFolder);
    }
}