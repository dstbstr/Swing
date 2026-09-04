 import { GetSingleRow} from "../Utils/SheetUtils"
 import { GetAttendanceSheetCurrentMonth, GetWaiverSheet, FindUserIndex, SheetDetails } from "../Utils/WoodsideUtils"

export default function CopyLatestWaiverToAttendance () {
    const [fullName, parent] = GetNewData();
    var attendanceSheet = GetAttendanceSheetCurrentMonth();
    UpdateAttendance(attendanceSheet, fullName, parent);
}

const GetNewData = (): [string, string] => {
    var waiverSheet = GetWaiverSheet();
    var latestRow = GetSingleRow(waiverSheet, waiverSheet.getLastRow());
    var sheetDetails = new SheetDetails(waiverSheet);
    var fullName = latestRow[sheetDetails.ParticipantColumn].trim();
    var signature = latestRow[sheetDetails.SignatureColumn].trim();
    var parent = latestRow[sheetDetails.ParentColumn].trim();

    if (!IsFullName(fullName) && IsFullName(signature) &&
        (parent === "" || signature.toLowerCase().startsWith(fullName.charAt(0).toLowerCase()))) {
        fullName = signature;
    }

    return [fullName, parent];
};

const IsFullName = (name: string): boolean => name.trim().split(/\s+/).length > 1;

const UpdateAttendance = (sheet: GoogleAppsScript.Spreadsheet.Sheet, fullName: string, parent: string) => {
    const sheetDetails = new SheetDetails(sheet);
    var existingIndex = FindUserIndex(sheet, fullName, sheetDetails.FullNameColumn);
    if (existingIndex === undefined) {
        var newRow = new Array(sheet.getLastColumn());
        newRow[sheetDetails.FullNameColumn] = fullName;
        newRow[sheetDetails.NotesColumn] = parent;
        sheet.appendRow(newRow);
        Logger.log(`Added new row for ${fullName}`);
    }
    else {
        Logger.log(`User ${fullName} already exists.`);
    }
};
