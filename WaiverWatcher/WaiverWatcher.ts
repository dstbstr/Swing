// seems like there is a bug in clasp which generates bad calls.
//solution: comment out the line below before running `clasp push`
import { GetSingleFolder, GetSingleFile, GetSingleSheet, GetSingleRow, IndexToHeader, FindColumnIndex} from "../Utils/SheetUtils.ts"
import {MONTHS, PARENT_FOLDER_NAME} from "../Utils/Constants.ts"

const FIRST_NAME_REGEX = /first/i;
const LAST_NAME_REGEX = /last/i;
const WAIVER_REGEX = /waiver/i;
const WAIVER_NOTES_REGEX = /know/i;
const ATTENDENCE_NOTES_REGEX = /notes/i;
const MINORS_REGEX = /minor/i;

export default function CopyLatestWaiverToAttendance () {
    const [firstName, lastName/*, notes, minors*/] = GetNewData();
    var attendenceSheet = GetAttendenceSheet();
    UpdateAttendence(attendenceSheet, firstName, lastName/*, notes*/);
    /*
    minors.forEach(m => {
        var [minorFirstName, minorLastName] = m.split(" ", 2);
        UpdateAttendence(attendenceSheet, minorFirstName, minorLastName, notes);
    });
    */
}

const GetNewData = (): [string, string/*, string, string[]*/] => {
    var waiverSheet = GetWaiverSheet();
    var latestRow = GetSingleRow(waiverSheet, waiverSheet.getLastRow());
    var lut = IndexToHeader(waiverSheet);
    var firstNameIdx = FindColumnIndex(lut, FIRST_NAME_REGEX);
    var lastNameIdx = FindColumnIndex(lut, LAST_NAME_REGEX);
    /*
    var notesIdx = FindColumnIndex(lut, WAIVER_NOTES_REGEX);
    var minorsIdx = FindColumnIndex(lut, MINORS_REGEX);
    var minors = latestRow[minorsIdx]
        .split("\n")
        .map(m => m.trim())
        .filter(m => m.length > 0);
        */
    //return [latestRow[firstNameIdx], latestRow[lastNameIdx], latestRow[notesIdx], minors];
    return [latestRow[firstNameIdx], latestRow[lastNameIdx]];
};

const GetAttendenceSheet = () : GoogleAppsScript.Spreadsheet.Sheet => {
    const currentYear = new Date().getFullYear();
    const currentMonth = new Date().getMonth();
    const parentFolder = GetSingleFolder(PARENT_FOLDER_NAME);
    //const attandenceFile = GetSingleFile(parentFolder, `Example Attendance ${currentYear}`);
    const attendanceFile = GetSingleFile(parentFolder, `Woodside Attendence ${currentYear}`);
    return GetSingleSheet(SpreadsheetApp.open(attendanceFile), MONTHS[currentMonth]);
};

const GetWaiverSheet = () : GoogleAppsScript.Spreadsheet.Sheet => {
    const currentYear = new Date().getFullYear();
    const parentFolder = GetSingleFolder(PARENT_FOLDER_NAME);
    //const waiverFile = GetSingleFile(parentFolder, `Example Waiver ${currentYear} (Responses)`);
    const waiverFile = GetSingleFile(parentFolder, `Woodside Waiver ${currentYear} Responses`);
    const spreadsheet = SpreadsheetApp.open(waiverFile);
    return spreadsheet.getActiveSheet();
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

const UpdateAttendence = (sheet: GoogleAppsScript.Spreadsheet.Sheet, firstName: string, lastName: string/*, notes: string*/) => {
    var lut = IndexToHeader(sheet);
    var firstNameIdx = FindColumnIndex(lut, FIRST_NAME_REGEX);
    var lastNameIdx = FindColumnIndex(lut, LAST_NAME_REGEX);
    var waiverIdx = FindColumnIndex(lut, WAIVER_REGEX);
    //var notesIdx = FindColumnIndex(lut, ATTENDENCE_NOTES_REGEX);
    var existingIndex = FindExistingIndex(sheet, firstName, lastName, firstNameIdx, lastNameIdx);
    if (existingIndex === undefined) {
        var newRow = new Array(sheet.getLastColumn());
        newRow[firstNameIdx] = firstName;
        newRow[lastNameIdx] = lastName;
        newRow[waiverIdx] = "Yes";
        //newRow[notesIdx] = notes;
        sheet.appendRow(newRow);
        Logger.log(`Added new row for ${firstName} ${lastName}`);
    }
    else {
        var existingRow = GetSingleRow(sheet, existingIndex);
        existingRow[waiverIdx] = "Yes";
        //existingRow[notesIdx] = notes;
        sheet.getRange(existingIndex, 1, 1, existingRow.length).setValues([existingRow]);
        Logger.log(`Updated row for ${firstName} ${lastName}`);
    }
};
