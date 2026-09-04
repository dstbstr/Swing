import {MONTHS, MONTHS_LONG, PARENT_FOLDER_NAME, PARTICIPANT_REGEX, PARENT_REGEX, FULL_NAME_REGEX, NOTES_REGEX, SIGNATURE_REGEX} from "../Utils/Constants"

import { GetSingleFolder, GetSingleFile, GetSingleSheet, IndexToHeader, FindColumnIndex, FindFirstDateIndex} from "../Utils/SheetUtils"

export const GetAttendanceFile = () : GoogleAppsScript.Spreadsheet.Spreadsheet => {
    const currentYear = new Date().getFullYear();
    const parentFolder = GetSingleFolder(PARENT_FOLDER_NAME);
    const file =  GetSingleFile(parentFolder, `Attendance ${currentYear}`);
    return SpreadsheetApp.open(file);
}

export const GetAttendanceSheetCurrentMonth = () : GoogleAppsScript.Spreadsheet.Sheet => {
    const currentMonth = new Date().getMonth();
    const file = GetAttendanceFile();
    return GetSingleSheet(file, MONTHS[currentMonth]);
};

export const GetWaiverSheet = () : GoogleAppsScript.Spreadsheet.Sheet => {
    const currentYear = new Date().getFullYear();
    const parentFolder = GetSingleFolder(PARENT_FOLDER_NAME);
    const waiverFile = GetSingleFile(parentFolder, `HCOS - Woodside Waiver ${currentYear} (Responses)`);
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
    const file = GetVolunteerFile();
    return GetSingleSheet(file, MONTHS_LONG[currentMonth]);
}

export const GetPracticeFile = () : GoogleAppsScript.Spreadsheet.Spreadsheet => {
    const currentYear = new Date().getFullYear();
    const parentFolder = GetSingleFolder(PARENT_FOLDER_NAME);
    const file = GetSingleFile(parentFolder, `Practice Night ${currentYear}`);
    return SpreadsheetApp.open(file);
}

export class SheetDetails {
    FullNameColumn: number;
    ParticipantColumn: number;
    ParentColumn: number;
    NotesColumn: number;
    SignatureColumn: number;
    FirstWeekColumn: number;
    Lut: { [key: string]: number };
    constructor(public sheet: GoogleAppsScript.Spreadsheet.Sheet, public caseInsensitive: boolean = true) {
        this.Lut = IndexToHeader(sheet, caseInsensitive);
        this.FullNameColumn = FindColumnIndex(this.Lut, FULL_NAME_REGEX) ?? -1;
        this.NotesColumn = FindColumnIndex(this.Lut, NOTES_REGEX) ?? -1;
        this.ParticipantColumn = FindColumnIndex(this.Lut, PARTICIPANT_REGEX) ?? -1;
        this.ParentColumn = FindColumnIndex(this.Lut, PARENT_REGEX) ?? -1;
        this.SignatureColumn = FindColumnIndex(this.Lut, SIGNATURE_REGEX) ?? -1;
        this.FirstWeekColumn = FindFirstDateIndex(this.Lut) ?? -1;;
    }
}
export const FindUserIndex = (sheet: GoogleAppsScript.Spreadsheet.Sheet, fullName: string, fullNameIdx: number): number | undefined => {
    var data = sheet.getDataRange().getValues();
    //skip header
    for (var row = 1; row < data.length; row++) {
        if (data[row][fullNameIdx].trim() === fullName.trim()) {
            return row + 1; //rows are 1 indexed
        }
    }
    return undefined;
}