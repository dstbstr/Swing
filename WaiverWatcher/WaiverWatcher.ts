// seems like there is a bug in clasp which generates bad calls.
//solution: comment out the lines below before running `clasp push`
 //import { GetSingleRow} from "../Utils/SheetUtils.ts"
 //import { GetAttendenceSheetCurrentMonth, GetWaiverSheet, FindUserIndex, SheetDetails } from "../Utils/WoodsideUtils.ts"


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
    var sheetDetails = new SheetDetails(waiverSheet);
    /*
    var notesIdx = FindColumnIndex(lut, WAIVER_NOTES_REGEX);
    var minorsIdx = FindColumnIndex(lut, MINORS_REGEX);
    var minors = latestRow[minorsIdx]
        .split("\n")
        .map(m => m.trim())
        .filter(m => m.length > 0);
        */
    //return [latestRow[firstNameIdx], latestRow[lastNameIdx], latestRow[notesIdx], minors];
    return [latestRow[sheetDetails.FirstNameColumn].trim(), latestRow[sheetDetails.LastNameColumn].trim()];
};

const UpdateAttendence = (sheet: GoogleAppsScript.Spreadsheet.Sheet, firstName: string, lastName: string/*, notes: string*/) => {
    const sheetDetails = new SheetDetails(sheet);
    //var notesIdx = FindColumnIndex(lut, ATTENDENCE_NOTES_REGEX);
    var filter = sheet.getFilter();
    var filterCriteria: GoogleAppsScript.Spreadsheet.FilterCriteria|null = null;
    if (filter !== null) {
        filterCriteria = filter.getColumnFilterCriteria(sheetDetails.WaiverColumn + 1);
        filter.remove();
    }
    var existingIndex = FindUserIndex(sheet, firstName, lastName, sheetDetails.FirstNameColumn, sheetDetails.LastNameColumn);
    if (existingIndex === undefined) {
        var newRow = new Array(sheet.getLastColumn());
        newRow[sheetDetails.FirstNameColumn] = firstName;
        newRow[sheetDetails.LastNameColumn] = lastName;
        newRow[sheetDetails.WaiverColumn] = "Yes";
        //newRow[sheetDetails.NotesColumn] = notes;
        sheet.appendRow(newRow);
        Logger.log(`Added new row for ${firstName} ${lastName}`);
    }
    else {
        var existingRow = GetSingleRow(sheet, existingIndex);
        //TODO: if waiver is already yes, could this be a duplicate?
        existingRow[sheetDetails.WaiverColumn] = "Yes";
        //existingRow[notesIdx] = notes;
        sheet.getRange(existingIndex, 1, 1, existingRow.length).setValues([existingRow]);
        Logger.log(`Updated row for ${firstName} ${lastName}`);
    }

    if(filterCriteria !== null) {
        var range = sheet.getDataRange();
        var newFilter = range.createFilter();
        newFilter.setColumnFilterCriteria(sheetDetails.WaiverColumn + 1, filterCriteria);
    }
};
