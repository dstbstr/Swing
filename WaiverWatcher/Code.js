// Compiled using undefined undefined (TypeScript 4.9.5)
var exports = exports || {};
var module = module || { exports: exports };
//import "./SheetUtils";
var FIRST_NAME_REGEX = /first ?name/i;
var LAST_NAME_REGEX = /last ?name/i;
var WAIVER_REGEX = /waiver/i;
var WAIVER_NOTES_REGEX = /know/i;
var ATTENDENCE_NOTES_REGEX = /notes/i;
var MINORS_REGEX = /minor/i;
var PARENT_FOLDER_NAME = "Scripting";
var NOT_FOUND = -1;
function OnSubmit() {
    var _a = GetNewData(), firstName = _a[0], lastName = _a[1], notes = _a[2], minors = _a[3];
    var attendenceSheet = GetAttendenceSheet();
    UpdateAttendence(attendenceSheet, firstName, lastName, notes);
    minors.forEach(function (minor) {
        var _a = minor.split(" ", 2), minorFirstName = _a[0], minorLastName = _a[1];
        UpdateAttendence(attendenceSheet, minorFirstName, minorLastName, notes);
    });
}
var MonthToName = function (month) {
    return ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][month];
};
var GetNewData = function () {
    var waiverSheet = SpreadsheetApp.getActiveSheet();
    var latestRow = SheetUtils.GetSingleRow(waiverSheet, waiverSheet.getLastRow());
    var lut = SheetUtils.IndexToHeader(waiverSheet);
    var firstNameIdx = SheetUtils.FindColumnIndex(lut, FIRST_NAME_REGEX);
    var lastNameIdx = SheetUtils.FindColumnIndex(lut, LAST_NAME_REGEX);
    var notesIdx = SheetUtils.FindColumnIndex(lut, WAIVER_NOTES_REGEX);
    var minorsIdx = SheetUtils.FindColumnIndex(lut, MINORS_REGEX);
    var minors = latestRow[minorsIdx]
        .split("\n")
        .map(function (m) { return m.trim(); })
        .filter(function (m) { return m.length > 0; });
    return [latestRow[firstNameIdx], latestRow[lastNameIdx], latestRow[notesIdx], minors];
};
var GetAttendenceSheet = function () {
    var currentYear = new Date().getFullYear();
    var currentMonth = new Date().getMonth();
    var parentFolder = SheetUtils.GetSingleFolder(PARENT_FOLDER_NAME);
    var attendenceFile = SheetUtils.GetSingleFile(parentFolder, "Example Attendance ".concat(currentYear));
    return SheetUtils.GetSingleSheet(SpreadsheetApp.open(attendenceFile), MonthToName(currentMonth));
};
var FindExistingIndex = function (sheet, firstName, lastName, lut, firstNameIdx, lastNameIdx) {
    var data = sheet.getDataRange().getValues();
    //skip header
    for (var row = 1; row < data.length; row++) {
        if (data[row][firstNameIdx] === firstName && data[row][lastNameIdx] === lastName) {
            return row + 1; //rows are 1 indexed
        }
    }
    return NOT_FOUND;
};
var UpdateAttendence = function (sheet, firstName, lastName, notes) {
    var lut = SheetUtils.IndexToHeader(sheet);
    var firstNameIdx = SheetUtils.FindColumnIndex(lut, FIRST_NAME_REGEX);
    var lastNameIdx = SheetUtils.FindColumnIndex(lut, LAST_NAME_REGEX);
    var waiverIdx = SheetUtils.FindColumnIndex(lut, WAIVER_REGEX);
    var notesIdx = SheetUtils.FindColumnIndex(lut, ATTENDENCE_NOTES_REGEX);
    var existingIndex = FindExistingIndex(sheet, firstName, lastName, lut, firstNameIdx, lastNameIdx);
    if (existingIndex == NOT_FOUND) {
        var newRow = new Array(sheet.getLastColumn());
        newRow[firstNameIdx] = firstName;
        newRow[lastNameIdx] = lastName;
        newRow[waiverIdx] = "Yes";
        newRow[notesIdx] = notes;
        sheet.appendRow(newRow);
        Logger.log("Added new row for ".concat(firstName, " ").concat(lastName));
    }
    else {
        var existingRow = SheetUtils.GetSingleRow(sheet, existingIndex);
        existingRow[waiverIdx] = "Yes";
        existingRow[notesIdx] = notes;
        sheet.getRange(existingIndex, 1, 1, existingRow.length).setValues([existingRow]);
        Logger.log("Updated row for ".concat(firstName, " ").concat(lastName));
    }
    sheet.setFrozenRows(1);
    sheet.sort(lastNameIdx);
};
