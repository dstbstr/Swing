// seems like there is a bug in clasp which generates bad calls.
//solution: comment out the lines below before running `clasp push`
// import { GetSingleRow, IndexToHeader, FindColumnIndex} from "../Utils/SheetUtils.ts"
// import { GetAttendenceSheetCurrentMonth, GetWaiverSheet } from "../Utils/WoodsideUtils.ts"
import { FIRST_NAME_REGEX, LAST_NAME_REGEX, WAIVER_REGEX } from "../Utils/Constants.ts"

//const WAIVER_NOTES_REGEX = /know/i;
//const ATTENDENCE_NOTES_REGEX = /notes/i;
//const MINORS_REGEX = /minor/i;

export default function CopyLatestWaiverToAttendance () {
    const [firstName, lastName/*, notes, minors*/] = GetNewData();
    var attendenceSheet = GetAttendenceSheetCurrentMonth();
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
    var filter = sheet.getFilter();
    var filterCriteria: GoogleAppsScript.Spreadsheet.FilterCriteria|null = null;
    if (filter !== null) {
        filterCriteria = filter.getColumnFilterCriteria(waiverIdx + 1);
        filter.remove();
    }
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

    if(filterCriteria !== null) {
        var range = sheet.getDataRange();
        var newFilter = range.createFilter();
        newFilter.setColumnFilterCriteria(waiverIdx + 1, filterCriteria);
    }
};
