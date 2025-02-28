// Compiled using undefined undefined (TypeScript 4.9.5)
var SheetUtils;
(function (SheetUtils) {
    SheetUtils.GetSingleFolder = function (folderName) {
        var parentFolders = DriveApp.getFoldersByName(folderName);
        if (parentFolders.hasNext() === false) {
            throw new Error("Could not find folder ".concat(folderName));
        }
        var result = parentFolders.next();
        if (parentFolders.hasNext()) {
            throw new Error("Duplicate folder found with name ".concat(folderName));
        }
        return result;
    };
    SheetUtils.GetSingleFile = function (parentFolder, fileName) {
        var targetFiles = parentFolder.getFilesByName(fileName);
        if (targetFiles.hasNext() === false) {
            throw new Error("Could not find file in folder ".concat(parentFolder.getName(), "/").concat(fileName));
        }
        var result = targetFiles.next();
        if (targetFiles.hasNext()) {
            throw new Error("Duplicate file found in folder with name ".concat(parentFolder.getName(), "/").concat(fileName));
        }
        return result;
    };
    SheetUtils.GetSingleSheet = function (sheet, sheetName) {
        var targetSheets = sheet.getSheets().filter(function (s) { return s.getName() === sheetName; });
        if (targetSheets.length != 1) {
            throw new Error("Could not find sheet ".concat(sheetName));
        }
        return targetSheets[0];
    };
    SheetUtils.GetSingleRow = function (sheet, row) {
        var rowRect = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues();
        if (rowRect.length != 1) {
            throw new Error("Row ".concat(row, " is not a single row"));
        }
        return rowRect[0];
    };
    SheetUtils.IndexToHeader = function (sheet) {
        var firstRow = SheetUtils.GetSingleRow(sheet, 1);
        var result = {};
        firstRow.forEach(function (value, index) { return result[value.toString().toLowerCase()] = index; });
        return result;
    };
    SheetUtils.FindColumnIndex = function (lut, pattern) {
        var result = Object.entries(lut).find(function (_a) {
            var key = _a[0], value = _a[1];
            return pattern.test(key);
        });
        if (result === undefined) {
            throw new Error("Could not find column with pattern ".concat(pattern));
        }
        return result[1];
    };
})(SheetUtils || (SheetUtils = {}));
