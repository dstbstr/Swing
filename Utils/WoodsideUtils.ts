import {MONTHS, PARENT_FOLDER_NAME} from "../Utils/Constants.ts"
//import { GetSingleFolder, GetSingleFile, GetSingleSheet} from "../Utils/SheetUtils.ts"

export const GetAttendenceFile = () : GoogleAppsScript.Spreadsheet.Spreadsheet => {
    const currentYear = new Date().getFullYear();
    const parentFolder = GetSingleFolder(PARENT_FOLDER_NAME);
    const file =  GetSingleFile(parentFolder, `Woodside Attendence ${currentYear}`);
    return SpreadsheetApp.open(file);
}

export const GetAttendenceSheetCurrentMonth = () : GoogleAppsScript.Spreadsheet.Sheet => {
    const currentMonth = new Date().getMonth();
    const file = GetAttendenceFile();
    return GetSingleSheet(file, MONTHS[currentMonth]);
};

export const GetWaiverSheet = () : GoogleAppsScript.Spreadsheet.Sheet => {
    const currentYear = new Date().getFullYear();
    const parentFolder = GetSingleFolder(PARENT_FOLDER_NAME);
    const waiverFile = GetSingleFile(parentFolder, `Woodside Waiver ${currentYear} Responses`);
    const spreadsheet = SpreadsheetApp.open(waiverFile);
    return spreadsheet.getActiveSheet();
};
