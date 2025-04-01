const GetSingleFolder = (folderName: string): GoogleAppsScript.Drive.Folder => {
    var parentFolders = DriveApp.getFoldersByName(folderName);
    if (parentFolders.hasNext() === false) {
        throw new Error(`Could not find folder ${folderName}`);
    }
    var result = parentFolders.next();
    if (parentFolders.hasNext()) {
        throw new Error(`Duplicate folder found with name ${folderName}`);
    }
    return result;
}
const GetSingleFile = (parentFolder: GoogleAppsScript.Drive.Folder, fileName: string): GoogleAppsScript.Drive.File => {
    var targetFiles = parentFolder.getFilesByName(fileName);
    if (targetFiles.hasNext() === false) {
        throw new Error(`Could not find file in folder ${parentFolder.getName()}/${fileName}`);
    }
    var result = targetFiles.next();
    if (targetFiles.hasNext()) {
        throw new Error(`Duplicate file found in folder with name ${parentFolder.getName()}/${fileName}`);
    }
    return result;
};
const TryGetSingleSheet = (sheet: GoogleAppsScript.Spreadsheet.Spreadsheet, sheetName: string): GoogleAppsScript.Spreadsheet.Sheet | undefined => {
    var targetSheets = sheet.getSheets().filter(function (s) { return s.getName() === sheetName; });
    if (targetSheets.length != 1) {
        return undefined;
    }
    return targetSheets[0];
};

const GetSingleSheet = (sheet: GoogleAppsScript.Spreadsheet.Spreadsheet, sheetName: string): GoogleAppsScript.Spreadsheet.Sheet => {
    var targetSheets = sheet.getSheets().filter(function (s) { return s.getName() === sheetName; });
    if (targetSheets.length != 1) {
        throw new Error("Could not find sheet ".concat(sheetName));
    }
    return targetSheets[0];
};
const GetSingleRow = (sheet: GoogleAppsScript.Spreadsheet.Sheet, row: number): any[] => {
    var rowRect = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues();
    if (rowRect.length != 1) {
        throw new Error(`Row ${row} is not a single row`);
    }
    return rowRect[0];
};
const IndexToHeader = (sheet: GoogleAppsScript.Spreadsheet.Sheet, caseInsensitive: boolean = true): { [key: string]: number } => {
    var firstRow = GetSingleRow(sheet, 1);
    var result = {};
    if(caseInsensitive) {
        firstRow.forEach((value, index) => result[value.toString().toLowerCase()] = index);
    } else {
        firstRow.forEach((value, index) => result[value.toString()] = index);
    }
    return result;
}
const FindColumnIndex = (lut: { [key: string]: number }, pattern: RegExp): number => {
    var result = Object.entries(lut).find(([key, value]) => pattern.test(key));
    if (result === undefined) {
        throw new Error(`Could not find column with pattern ${pattern}`);
    }
    return result[1];
};
const FindFirstDateIndex = (lut: { [key: string]: number }): number => {
    var result = Object.entries(lut).find(([key, value]) => {
        var date = new Date(key);
        return !isNaN(date.getTime());
    });
    if (result === undefined) {
        throw new Error(`Could not find first date column`);
    }
    return result[1];
}

export {
    GetSingleFolder,
    GetSingleFile,
    TryGetSingleSheet,
    GetSingleSheet,
    GetSingleRow,
    IndexToHeader,
    FindColumnIndex,
    FindFirstDateIndex
}