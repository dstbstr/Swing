import {MONTHS} from "../Utils/Constants.ts"
//import {FindFirstDateIndex, IndexToHeader, TryGetSingleSheet} from "../Utils/SheetUtils.ts"
//import {GetAttendenceFile} from "../Utils/WoodsideUtils.ts"

const THURSDAY = 4;

export default function EnsureMonthExists() {
    var file = GetAttendenceFile();
    var month = new Date().getMonth();
    const currentMonth = TryGetSingleSheet(file, MONTHS[month]);
    const previousMonth = TryGetSingleSheet(file, MONTHS[month - 1]);
    if(currentMonth !== undefined) {
        Logger.log(`Sheet ${MONTHS[month]} already exists`);
        return;
    }
    if(previousMonth === undefined) {
        Logger.log(`Could not find ${MONTHS[month - 1]} sheet`);
        return;
    }

    Logger.log(`Creating ${MONTHS[month]} sheet`);

    var newSheet = file.insertSheet(MONTHS[month], file.getNumSheets());
    const nonDateHeaders = GetNonDateHeaders(previousMonth);
    const dateHeaders = GetDateHeaders();

    const headerValues = nonDateHeaders.concat(dateHeaders);

    newSheet.appendRow(headerValues);
    var headerRow = newSheet.getRange(1, 1, 1, headerValues.length);
    headerRow.setFontWeight("bold");
    newSheet.setFrozenRows(1);

    CopyPreviousMonth(newSheet, previousMonth, nonDateHeaders.length);

    const startLetter = String.fromCharCode('A'.charCodeAt(0) + nonDateHeaders.length);
    const endLetter = String.fromCharCode(startLetter.charCodeAt(0) + dateHeaders.length - 1);
    const startRow = 2;
    const endRow = newSheet.getDataRange().getNumRows() + 20;
    const dataRange = `${startLetter}${startRow}:${endLetter}${endRow}`;
    Logger.log(`Adding dropdowns to ${dataRange}`);

    AddDropdowns(newSheet.getRange(dataRange));
    Logger.log(`Created ${MONTHS[month]} sheet`);
}

const GetNonDateHeaders = (sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] => {
    const caseInsensitive = false;
    const lut = IndexToHeader(sheet, caseInsensitive);
    const firstDateHeader = FindFirstDateIndex(lut);
    let result: string[] = [];
    Object.entries(lut).forEach(([header, index], _) => {
        if (index < firstDateHeader) {
            result[index] = header;
        }
    });

    return result;    
}

const GetDateHeaders = (): string[] => {
    const date = new Date();
    const currentYear = date.getFullYear();
    const currentMonth = date.getMonth();
    const daysInMonth = new Date(currentYear, currentMonth, 0).getDate();
    const firstDayOfMonth = new Date(currentYear, currentMonth, 1).getDay();
    const firstThursday = ((11 - firstDayOfMonth) % 7) + 1;
    let result: string[] = [];
    
    for (var day = firstThursday; day < daysInMonth; day += 7) {
        const month = `${currentMonth + 1}`.padStart(2, "0");
        const dayStr = `${day}`.padStart(2, "0");
        result.push(`${month}-${dayStr}-${currentYear % 100}`);
    }
    return result;
}

const CopyPreviousMonth = (current: GoogleAppsScript.Spreadsheet.Sheet, previous: GoogleAppsScript.Spreadsheet.Sheet, headerCount: number) => {
    var previousData = previous.getDataRange().getValues();
    previousData.slice(1).forEach(function (row) {
        current.appendRow(row.slice(0, headerCount));
    });
}

const AddDropdowns = (range: GoogleAppsScript.Spreadsheet.Range) => {
    const dropdowns = SpreadsheetApp.newDataValidation()
        .requireValueInList([
            "",
            "Preregistered",
            "10cc",
            "10cash",
            "5cc+vou",
            "5cash+vou",
            "5cc+student",
            "5cash+student",
            "Student+vou",
            "5cash",
            "5cc",
            "vou",
            "volunteer"
        ])
        .setAllowInvalid(true)
        .build();

    range.setDataValidation(dropdowns);
}