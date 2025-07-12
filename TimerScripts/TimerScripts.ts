import {MONTHS, MONTHS_LONG} from "../Utils/Constants.ts"

// import {TryGetSingleSheet} from "../Utils/SheetUtils.ts"
// import {GetAttendenceFile, FindUserIndexByFullName, SheetDetails, GetVolunteerFile, GetPreregisterSheet} from "../Utils/WoodsideUtils.ts"

const THURSDAY = 4;

export default function EnsureThisMonth() {
    EnsureMonthExists(new Date().getMonth());
}
export const EnsureNextMonth = () => {
    EnsureMonthExists(new Date().getMonth() + 1);
}
export const UpdateVolunteers = () => {
    const targetMonth = new Date().getMonth();
    const sheet = GetAttendenceSheetCurrentMonth();
    const dateHeaders = GetDateHeaders(targetMonth);
    HighlightVolunteers(sheet, targetMonth, dateHeaders.length);
}

const EnsureMonthExists = (targetMonth: number) => {
    var file = GetAttendenceFile();
    const currentMonth = TryGetSingleSheet(file, MONTHS[targetMonth]);
    const previousMonth = TryGetSingleSheet(file, MONTHS[targetMonth - 1]);
    if(currentMonth !== undefined) {
        Logger.log(`Sheet ${MONTHS[targetMonth]} already exists`);
        return;
    }
    if(previousMonth === undefined) {
        Logger.log(`Could not find ${MONTHS[targetMonth - 1]} sheet`);
        return;
    }

    CreateMonth(targetMonth, file, previousMonth);
}

const CreateMonth = (month: number, file: GoogleAppsScript.Spreadsheet.Spreadsheet, prevMonth: GoogleAppsScript.Spreadsheet.Sheet) => {
    Logger.log(`Creating ${MONTHS[month]} sheet`);

    var newSheet = file.insertSheet(MONTHS[month], file.getNumSheets());
    const nonDateHeaders = GetNonDateHeaders(prevMonth);
    const dateHeaders = GetDateHeaders(month);

    const headerValues = nonDateHeaders.concat(dateHeaders);

    newSheet.appendRow(headerValues);
    var headerRow = newSheet.getRange(1, 1, 1, headerValues.length);
    headerRow.setFontWeight("bold");
    newSheet.setFrozenRows(1);
    headerRow.protect().setWarningOnly(true);

    CopyPreviousMonth(newSheet, prevMonth, nonDateHeaders.length);
    HighlightVolunteers(newSheet, month, dateHeaders.length);

    const startLetter = String.fromCharCode('A'.charCodeAt(0) + nonDateHeaders.length);
    const endLetter = String.fromCharCode(startLetter.charCodeAt(0) + dateHeaders.length - 1);
    const startRow = 2;
    const endRow = newSheet.getDataRange().getNumRows() + 20;
    const dataRange = `${startLetter}${startRow}:${endLetter}${endRow}`;
    Logger.log(`Adding dropdowns to ${dataRange}`);

    AddDropdowns(newSheet.getRange(dataRange));
    Logger.log(`Created ${MONTHS[month]} sheet`);
};

const GetNonDateHeaders = (sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] => {
    const sheetDetails = new SheetDetails(sheet, false);
    let result: string[] = [];
    Object.entries(sheetDetails.Lut).forEach(([header, index], _) => {
        if (index < sheetDetails.FirstWeekColumn) {
            result[index] = header;
        }
    });

    return result;    
}

const GetDateHeaders = (targetMonth: number): string[] => {
    const date = new Date();
    const currentYear = date.getFullYear();
    const daysInMonth = new Date(currentYear, targetMonth, 0).getDate();
    const firstDayOfMonth = new Date(currentYear, targetMonth, 1).getDay();
    const firstThursday = ((11 - firstDayOfMonth) % 7) + 1;
    let result: string[] = [];
    
    for (var day = firstThursday; day < daysInMonth; day += 7) {
        const month = `${targetMonth + 1}`.padStart(2, "0");
        const dayStr = `${day}`.padStart(2, "0");
        result.push(`${month}-${dayStr}-${currentYear % 100}`);
    }
    return result;
}

const CopyPreviousMonth = (current: GoogleAppsScript.Spreadsheet.Sheet, previous: GoogleAppsScript.Spreadsheet.Sheet, headerCount: number) => {
    const previousData = previous.getDataRange().getValues();
    const sortColumn = 0; // meh, could look it up, but would require parsing the sheet details
    const dataToCopy = previousData
        .slice(1) // skip header row
        .map(row => row.slice(0, headerCount)) // grab everything before the dates
        .map(row => row.map(cell => cell.toString().trim())) // trim the cells
        .filter(row => row[0] !== "" && row[1] !== "") // remove empty names
        .sort((a, b) => { return a[sortColumn].localeCompare(b[sortColumn]); });
    if(dataToCopy.length === 0) return;
    current.getRange(2, 1, dataToCopy.length, headerCount).setValues(dataToCopy);
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
            "volunteer",
            "promotion"
        ])
        .setAllowInvalid(true)
        .build();

    range.setDataValidation(dropdowns);
}

const AddEventDropdowns = (range: GoogleAppsScript.Spreadsheet.Range) => {
    const dropdowns = SpreadsheetApp.newDataValidation()
        .requireValueInList([
            "",
            "Preregistered",
            "20cc",
            "20cash",
            "15cc+vou",
            "15cash+vou",
            "15cc+student",
            "15cash+student",
            "10cc+vou+student",
            "10cash+vou+student",
            "10cc+volunteer",
            "10cash+volunteer"
        ])
        .setAllowInvalid(true)
        .build();
    range.setDataValidation(dropdowns);
}

const AddMondayDropdowns = (range: GoogleAppsScript.Spreadsheet.Range) => {
    const dropdowns = SpreadsheetApp.newDataValidation()
        .requireValueInList([
            "",
            "Preregistered",
            "5cc",
            "5cash",
            "vou",
            "volunteer",
            "promotion"
        ])
        .setAllowInvalid(true)
        .build();
    range.setDataValidation(dropdowns);
}

const HighlightVolunteers = (sheet: GoogleAppsScript.Spreadsheet.Sheet, targetMonth: number, numDates: number) => {
    const LEADERS = [
        "Darcy Brown",
        "Amanda Darr",
        "Jason Goetz",
        "Tom Hamming",
        "Dustin Randall"
    ];
            
    const sheetDetails = new SheetDetails(sheet);
    const dateIdxs: {[key: string]: number} = {};
    Object.entries(sheetDetails.Lut).forEach(([header, index], _) => {
        let date = new Date(header);
        if(date !== undefined) {
            dateIdxs[FormatDate(date)] = index;
        }
    });

    LEADERS.forEach(leader => {
        const row = FindUserIndexByFullName(sheet, leader, sheetDetails.FirstNameColumn, sheetDetails.LastNameColumn);
        if(row === undefined) {
            Logger.log(`Could not find row for ${leader}`);
            return;
        }
        const rowRange = sheet.getRange(row, sheetDetails.FirstWeekColumn + 1, 1, numDates);
        rowRange.setBackground("yellow");
    });

    const volunteers = GetVolunteers(targetMonth);
    Object.entries(volunteers).forEach(([date, names]) => {
        const dateIdx = dateIdxs[date];
        if(dateIdx === undefined) {
            Logger.log(`Could not find date ${date} in attendance sheet`);
            return;
        }
        names.forEach(name => {
            const row = FindUserIndexByFullName(sheet, name, sheetDetails.FirstNameColumn, sheetDetails.LastNameColumn);
            if(row === undefined) {
                Logger.log(`Could not find ${name} in attendance sheet`);
                return;
            }

            const rowRange = sheet.getRange(row, dateIdx + 1);
            rowRange.setBackground("yellow");
        });
    });

    const preregs = GetPreregs(targetMonth)
    preregs.forEach(name => {
        const row = FindUserIndexByFullName(sheet, name, sheetDetails.FirstNameColumn, sheetDetails.LastNameColumn);
        if(row === undefined) {
            Logger.log(`Could not find ${name} in attendance sheet`);
            return;
        }
        const rowRange = sheet.getRange(row, sheetDetails.FirstWeekColumn + 1, 1, numDates);
        rowRange.setBackground("green");
    });
}

const GetVolunteers = (targetMonth: number): {[key: string]: string[]} => {
    const file = GetVolunteerFile();
    const sheet = TryGetSingleSheet(file, MONTHS_LONG[targetMonth]);
    var result = {};
    if(sheet === undefined) {
        Logger.log(`Could not find ${MONTHS_LONG[targetMonth]} sheet`);
        return result;
    }

    const data = sheet.getDataRange().getValues();
    for (var row = 1; row < data.length; row++) {
        var dateStr = data[row][0];
        var date = new Date(dateStr);
        if(isNaN(date.getTime())) continue;
        if(date.getDay() !== THURSDAY) continue; // Thursday only in this sheet

        dateStr = FormatDate(date);

        for(var col = 1; col < data[row].length; col++) {
            const name = data[row][col];
            if(name === "") continue;
            if(result[dateStr] === undefined) {
                result[dateStr] = [];
            }
            result[dateStr].push(name);
        }
        if(result[dateStr] === undefined) {
            Logger.log(`No volunteers found for ${dateStr}`);
        } else {
            Logger.log(`Found ${result[dateStr].length} volunteers for ${dateStr}`);
        }
    }

    return result;
}

const GetPreregs = (targetMonth: number): string[] => {
    const sheet = GetPreregisterSheet();
    const data = sheet.getDataRange().getValues();
    var result: string[] = [];
    var targetMonthName = MONTHS_LONG[targetMonth].toLowerCase();
    var nextMonthName = MONTHS_LONG[targetMonth + 1].toLowerCase();

    var inTargetMonth = false;

    for(var row = 1; row < data.length; row++) {
        const val = data[row][0].toString();
        if(val.toLowerCase().startsWith(targetMonthName)) {
            inTargetMonth = true;
            Logger.log(`Found target month ${targetMonthName} in Preregister sheet`);
            continue;
        }
        if(inTargetMonth) {
            // TODO: Brittle, expects each month to end with a blank line
            // Could check for next month name, except people might be named April, May, or June
            // Could check for bold/underline, or ending in 'Lindy Hop' but all of these are equally brittle
            if(val === "") break;
            result.push(val);
        }
    }
    return result;
}

const FormatDate = (date: Date): string => {
    var month = `${date.getMonth() + 1}`.padStart(2, "0");
    var day = `${date.getDate()}`.padStart(2, "0");
    return `${month}-${day}-${date.getFullYear() % 100}`;
}