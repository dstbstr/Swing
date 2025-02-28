import * as SheetUtils from "../Utils/SheetUtils.ts"
import {MONTHS, PARENT_FOLDER_NAME} from "../Utils/Constants.ts"

const FIRST_NAME_REGEX = /first ?name/i;
const LAST_NAME_REGEX = /last ?name/i;
const WAIVER_REGEX = /waiver/i;
const WAIVER_NOTES_REGEX = /know/i;
const ATTENDENCE_NOTES_REGEX = /notes/i;
const MINORS_REGEX = /minor/i;

const OnSubmit = () => {
    const [firstName, lastName, notes, minors] = GetNewData();
    var attendenceSheet = GetAttendenceSheet();
    UpdateAttendence(attendenceSheet, firstName, lastName, notes);
    minors.forEach(m => {
        var [minorFirstName, minorLastName] = m.split(" ", 2);
        UpdateAttendence(attendenceSheet, minorFirstName, minorLastName, notes);
    });
}

const GetNewData = (): [string, string, string, string[]] => {
    var waiverSheet = SpreadsheetApp.getActiveSheet();
    var latestRow = SheetUtils.GetSingleRow(waiverSheet, waiverSheet.getLastRow());
    var lut = SheetUtils.IndexToHeader(waiverSheet);
    var firstNameIdx = SheetUtils.FindColumnIndex(lut, FIRST_NAME_REGEX);
    var lastNameIdx = SheetUtils.FindColumnIndex(lut, LAST_NAME_REGEX);
    var notesIdx = SheetUtils.FindColumnIndex(lut, WAIVER_NOTES_REGEX);
    var minorsIdx = SheetUtils.FindColumnIndex(lut, MINORS_REGEX);
    var minors = latestRow[minorsIdx]
        .split("\n")
        .map(m => m.trim())
        .filter(m => m.length > 0);
    return [latestRow[firstNameIdx], latestRow[lastNameIdx], latestRow[notesIdx], minors];
};

const GetAttendenceSheet = () : GoogleAppsScript.Spreadsheet.Sheet => {
    const currentYear = new Date().getFullYear();
    const currentMonth = new Date().getMonth();
    const parentFolder = SheetUtils.GetSingleFolder(PARENT_FOLDER_NAME);
    const attendenceFile = SheetUtils.GetSingleFile(parentFolder, `Example Attendance ${currentYear}`);
    return SheetUtils.GetSingleSheet(SpreadsheetApp.open(attendenceFile), MONTHS[currentMonth]);
};

const FindExistingIndex = (sheet: GoogleAppsScript.Spreadsheet.Sheet, firstName: string, lastName: string, firstNameIdx: number, lastNameIdx: number) : number | undefined => {
    var data = sheet.getDataRange().getValues();
    //skip header
    for (var row = 1; row < data.length; row++) {
        if (data[row][firstNameIdx] === firstName && data[row][lastNameIdx] === lastName) {
            return row + 1; //rows are 1 indexed
        }
    }
    return undefined;
};

const UpdateAttendence = (sheet: GoogleAppsScript.Spreadsheet.Sheet, firstName: string, lastName: string, notes: string) => {
    var lut = SheetUtils.IndexToHeader(sheet);
    var firstNameIdx = SheetUtils.FindColumnIndex(lut, FIRST_NAME_REGEX);
    var lastNameIdx = SheetUtils.FindColumnIndex(lut, LAST_NAME_REGEX);
    var waiverIdx = SheetUtils.FindColumnIndex(lut, WAIVER_REGEX);
    var notesIdx = SheetUtils.FindColumnIndex(lut, ATTENDENCE_NOTES_REGEX);
    var existingIndex = FindExistingIndex(sheet, firstName, lastName, firstNameIdx, lastNameIdx);
    if (existingIndex === undefined) {
        var newRow = new Array(sheet.getLastColumn());
        newRow[firstNameIdx] = firstName;
        newRow[lastNameIdx] = lastName;
        newRow[waiverIdx] = "Yes";
        newRow[notesIdx] = notes;
        sheet.appendRow(newRow);
        Logger.log(`Added new row for ${firstName} ${lastName}`);
    }
    else {
        var existingRow = SheetUtils.GetSingleRow(sheet, existingIndex);
        existingRow[waiverIdx] = "Yes";
        existingRow[notesIdx] = notes;
        sheet.getRange(existingIndex, 1, 1, existingRow.length).setValues([existingRow]);
        Logger.log(`Updated row for ${firstName} ${lastName}`);
    }
    sheet.setFrozenRows(1);
    sheet.sort(lastNameIdx);
};
