import {MONTHS, MONTHS_LONG, PARENT_FOLDER_NAME, FIRST_NAME_REGEX, LAST_NAME_REGEX, WAIVER_REGEX} from "../Utils/Constants.ts"
// import { GetSingleFolder, GetSingleFile, GetSingleSheet, IndexToHeader, FindColumnIndex, FindFirstDateIndex} from "../Utils/SheetUtils.ts"

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

export const GetVolunteerFile = () : GoogleAppsScript.Spreadsheet.Spreadsheet => {
    const currentYear = new Date().getFullYear();
    const parentFolder = GetSingleFolder(PARENT_FOLDER_NAME);
    const file = GetSingleFile(parentFolder, `Volunteer Schedule ${currentYear}`);
    return SpreadsheetApp.open(file);
}
export const GetVolunteerSheetCurrentMonth = () : GoogleAppsScript.Spreadsheet.Sheet => {
    const currentMonth = new Date().getMonth();
    const file = GetAttendenceFile();
    return GetSingleSheet(file, MONTHS_LONG[currentMonth]);
}

export const GetPreregisterSheet = () : GoogleAppsScript.Spreadsheet.Sheet => {
    const parentFolder = GetSingleFolder(PARENT_FOLDER_NAME);
    const file = GetSingleFile(parentFolder, 'Woodside_Class Registration');
    return SpreadsheetApp.open(file).getActiveSheet();
}

export class SheetDetails {
    FirstNameColumn: number;
    LastNameColumn: number;
    FirstWeekColumn: number;
    Lut: { [key: string]: number };
    constructor(public sheet: GoogleAppsScript.Spreadsheet.Sheet, public caseInsensitive: boolean = true) {
        this.Lut = IndexToHeader(sheet, caseInsensitive);
        this.FirstNameColumn = FindColumnIndex(this.Lut, FIRST_NAME_REGEX) ?? -1;
        this.LastNameColumn = FindColumnIndex(this.Lut, LAST_NAME_REGEX) ?? -1;
        this.FirstWeekColumn = FindFirstDateIndex(this.Lut) ?? -1;;
    }
}
export const FindUserIndex = (sheet: GoogleAppsScript.Spreadsheet.Sheet, firstName: string, lastName: string, firstNameIdx: number, lastNameIdx: number): number | undefined => {
    var data = sheet.getDataRange().getValues();
    //skip header
    for (var row = 1; row < data.length; row++) {
        if (data[row][firstNameIdx].trim() === firstName.trim() && data[row][lastNameIdx].trim() === lastName.trim()) {
            return row + 1; //rows are 1 indexed
        }
    }
    return undefined;
}

export const FindUserIndexByFullName = (sheet: GoogleAppsScript.Spreadsheet.Sheet, name: string, firstNameIdx: number, lastNameIdx: number): number | undefined => {
    var split = name.split(" ", 2);
    if (split.length != 2) {
        return undefined;
    }
    return FindUserIndex(sheet, split[0], split[1], firstNameIdx, lastNameIdx);
}