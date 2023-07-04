/**
 * Encapsulation of miscallenous functions
 * Purpose : To ensure that there is no wrong usage of functions (Bypass)
 */
const miscTools = class {
    /**
     * Recursive Function to iterate every folders within Sub Directories and create new folders
     * @params {Object} args: Sub Directories that contains the parent folder,direct sub directories and nested sub directories
     * @params {Array} folderArray: The list of parent folders
     * Does not return any values
     * Does not throw any errors due to recursion
     */
    static recursiveFolderCreation(args, folderArray) {
        if (!args.length)
            return;
        for (const data of args) {
            let childName = data.parentFolder;
            let childSubFolders = data.subDirectories;
            let nestedData = data.args;
            let folder = folderArray.find(folder => folder.getName() === childName);
            if (!folder)
                continue;
            var childArray = [];
            childSubFolders.forEach(name => {
                let child = folderGenerator.retrieveFolderByName(name, folder);
                (!child) ?
                    childArray.push(folder.createFolder(name)) :
                    childArray.push(child);
            });
            if (!nestedData)
                continue;
            this.recursiveFolderCreation(nestedData, childArray);
        }
    }
    /**
     * Recursive Function to iterate every folders within Sub Directories
     * @params {Folder} parentFolder: The Parent Folder
     * @return {object} folderHierarchy: A nested JSON Object
     * Example: {parentFolder: rootFolderID, subDirectories: [folder1ID,folder2ID,folder3ID], args: {parentFolder : folder1ID, subDirectories: [folder4ID], args: {...args}}}
     */
    static recursiveFolder(parentFolder) {
        if (!parentFolder)
            return;
        let folderHierarchy = { id: parentFolder.getId(), name: parentFolder.getName(), subDirectories: [] };
        let subFolderIterator = parentFolder.getFolders();
        while (subFolderIterator.hasNext()) {
            let subFolder = subFolderIterator.next();
            folderHierarchy.subDirectories.push(this.recursiveFolder(subFolder));
        }
        return folderHierarchy;
    }

    /**
     * Recursive Function to iterate every folders to find file with specific file name
     * @params {string} fileName: Name of the file
     * @params {Folder} parentFolder: The root Folder which is being searched
     * @return {File} null if nothing is found, else return the file
     */
    static recursiveSearchFile(fileName, parentFolder) {
        if (!parentFolder)
            return null;
        let fileIterator = parentFolder.getFilesByName(fileName);
        while (fileIterator.hasNext())
            return fileIterator.next();
        let subFolderIterator = parentFolder.getFolders();
        while (subFolderIterator.hasNext()) {
            let subFolder = subFolderIterator.next();
            let file = this.recursiveSearchFile(fileName, subFolder);
            if (file)
                return file;
        }
        return null;
    }

    /**
     * Recursive Function to iterate every folders to find folder with specific folder name
     * @params {string} folderName: Name of the folder
     * @params {Folder} parentFolder: The root Folder which is being searched
     * @return {Folder} null if nothing is found, else return the folder
     */
    static recursiveSearchFolder(folderName, parentFolder) {
        if (!parentFolder)
            return null;
        if (String(parentFolder.getName()) === String(folderName))
            return parentFolder;
        let subFolderIterator = parentFolder.getFolders();
        while (subFolderIterator.hasNext()) {
            let subFolder = subFolderIterator.next();
            let folder = this.recursiveSearchFolder(folderName, subFolder);
            if (folder)
                return folder;
        }
        return null;
    }
}

