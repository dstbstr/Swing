const THURSDAY = 4;
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const EnsureMonthExists = () => {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var currentMonth = new Date().getMonth();
    var monthToCreate = MONTHS[currentMonth];
    var targetSheet = spreadsheet.getSheets().filter(function (s) { return s.getName() === monthToCreate; });
    if (targetSheet.length > 0) {
        Logger.log("Sheet ".concat(monthToCreate, " already exists"));
        return;
    }
    var newSheet = spreadsheet.insertSheet(monthToCreate, currentMonth);
    var headerValues = GetHeaderNames();
    newSheet.appendRow(headerValues);
    var headerRow = newSheet.getRange(1, 1, 1, headerValues.length);
    headerRow.setFontWeight("bold");
    newSheet.setFrozenRows(1);
    var previousSheet = GetPreviousSheet(spreadsheet, currentMonth);
    if (previousSheet === undefined) {
        Logger.log("Could not find previous month");
        return;
    }
    else {
        CopyPreviousMonth(newSheet, previousSheet);
    }
};

const GetHeaderNames = (): string[] => {
    var date = new Date();
    var currentYear = date.getFullYear();
    var currentMonth = date.getMonth();
    var daysInMonth = new Date(currentYear, currentMonth, 0).getDate();
    var firstDayOfMonth = new Date(currentYear, currentMonth, 1).getDay();
    var firstThursday = ((11 - firstDayOfMonth) % 7) + 1;
    var result = ["First Name", "Last Name", "Notes", "Waiver"];
    for (var day = firstThursday; day <= daysInMonth; day += 7) {
        const month = `${currentMonth + 1}`.padStart(2, "0");
        const dayStr = `${day}`.padStart(2, "0");
        result.push(`${month}-${dayStr}-${currentYear % 100}`);
    }
    return result;
};

const GetPreviousSheet = (sheet, currentMonth): GoogleAppsScript.Spreadsheet.Sheet | undefined => {
    if (currentMonth < 1) {
        return undefined;
    }
    var previousMonth = MONTHS[currentMonth - 1];
    var result = sheet.getSheets().filter(function (s) { return s.getName() === previousMonth; });
    if (result.length === 0) {
        return undefined;
    }
    return result[0];
};

const CopyPreviousMonth = (current, previous) => {
    var previousData = previous.getDataRange().getValues();
    previousData.slice(1).forEach(function (row) {
        current.appendRow(row.slice(0, 4));
    });
};
