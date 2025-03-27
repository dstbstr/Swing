// import { IndexToHeader, FindColumnIndex, FindFirstDateIndex, TryGetSingleSheet } from "../Utils/SheetUtils.ts"
// import { GetAttendenceFile } from "../Utils/WoodsideUtils.ts"
import { FIRST_NAME_REGEX, LAST_NAME_REGEX, WAIVER_REGEX, MONTHS } from "../Utils/Constants.ts"

export default function SendMonthlyReport() {
    const file = GetAttendenceFile();
    var monthStats: MonthStats[] = [];
    var summaries: RowSummary[][] = [];
    for (var i = 0; i < MONTHS.length; i++) {
        var sheet = TryGetSingleSheet(file, MONTHS[i]);
        if(sheet === undefined) continue;

        const lut = IndexToHeader(sheet);
        const firstNameColumn = FindColumnIndex(lut, FIRST_NAME_REGEX);
        const lastNameColumn = FindColumnIndex(lut, LAST_NAME_REGEX);
        const waiverColumn = FindColumnIndex(lut, WAIVER_REGEX);
        const firstWeekColumn = FindFirstDateIndex(lut);

        const sheetSummary = SummarizeSheet(sheet, firstNameColumn, lastNameColumn, waiverColumn, firstWeekColumn);
        const monthStat = SummarizeMonth(sheetSummary);
        monthStats[i] = monthStat; // keep month lined up with monthStats
    }
    const yearStats = SummarizeYear(monthStats);
    SendEmail(monthStats, yearStats);
}

class RowSummary {
    hasWaiver: boolean;
    fullName: string;
    hasFullName: boolean;
    attended: boolean[];
}

class MonthStats {
    countByWeek: number[];
    countByVisit: number[];
    missingWaivers: string[];
    duplicateNames: string[];
    incompleteNames: number;
    uniqueNames: Map<string, number>;
}

class YearStats {
    countByWeek: number[];
    countByVisit: number[];
    uniqueNames: Map<string, number>;
    uniqueNameCounts: number[];
}

const SummarizeSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet, firstNameColumn: number, lastNameColumn: number, waiverColumn: number, firstWeekColumn: number): RowSummary[] => {
    var result: RowSummary[] = [];
    var data = sheet.getDataRange().getValues();
    data.forEach((row, index) => {
        if (index === 0) { //skip header
            return;
        }
        var rowSummary: RowSummary = {
            hasWaiver: row[waiverColumn] !== "",
            fullName: `${row[firstNameColumn]} ${row[lastNameColumn]}`.trim(),
            hasFullName: row[firstNameColumn] !== "" && row[lastNameColumn] !== "",
            attended: []
        };
        for (var col = firstWeekColumn; col < row.length; col++) {
            rowSummary.attended.push(row[col] !== "");
        }
        if(rowSummary.hasWaiver === false && !rowSummary.attended.some(v => v)) {
            // only include rows that have attended or have signed a waiver this year
            return;
        }
        result.push(rowSummary);
    })

    return result;
}

const SummarizeMonth = (summary: RowSummary[]): MonthStats => {
    let result: MonthStats = {
        countByWeek: [],
        countByVisit: [],
        missingWaivers: [],
        duplicateNames: [],
        incompleteNames: 0,
        uniqueNames: new Map<string, number>()
    }

    summary.forEach(row => {
        const count = row.attended.reduce((sum, attended) => sum + (attended ? 1 : 0), 0);
        if(result.uniqueNames.has(row.fullName)) {
            result.duplicateNames.push(row.fullName);
        }
        result.uniqueNames.set(row.fullName, count);
        result.countByVisit[count] = (result.countByVisit[count] || 0) + 1;
        if(!row.hasWaiver) {
            result.missingWaivers.push(row.fullName);
        }
        result.incompleteNames += row.hasFullName ? 0 : 1;
    });

    result.countByWeek = CountByWeek(summary);

    return result;
    
}

const SummarizeYear = (summaries: MonthStats[]) : YearStats => {
    let result: YearStats = {
        countByWeek: [],
        countByVisit: [],
        uniqueNames: new Map<string, number>(),
        uniqueNameCounts: []
    };

    for(let i = 0; i < summaries.length; i++) {
        const summary = summaries[i];
        if(summary === undefined) continue;

        for(let week = 0; week < summary.countByWeek.length; week++) {
            result.countByWeek.push(summary.countByWeek[week]);
        }
        summary.uniqueNames.forEach((count, name) => {
            result.uniqueNames.set(name, (result.uniqueNames.get(name) || 0) + count);
        });
        result.uniqueNameCounts.push(result.uniqueNames.size);
    }
    result.uniqueNames.forEach((count, name) => {
        result.countByVisit[count] = (result.countByVisit[count] || 0) + 1;
    })

    return result;
}

const CountByWeek = (summary: RowSummary[]): number[] => {
    return summary.reduce((result, row) => {
        row.attended.forEach((present, weekIndex) => {
            result[weekIndex] = (result[weekIndex] || 0) + (present ? 1 : 0);
        });
        return result;
    }, [] as number[]);
}

const SendEmail = (monthStats: MonthStats[], yearStats: YearStats) => {
    const lastMonth = monthStats
        .filter(stats => stats !== undefined)
        .filter(stats => stats.uniqueNames.size > 0)
        .slice(-1)[0];
    if(lastMonth === undefined) {
        Logger.log("Nothing to send");
        return;
    }

    const cumulativeCountChart = CreateCumulativeCountChart(yearStats);
    const cumulativeVisitChart = CreateCumulativeVisitChart(yearStats);
    const countByWeekChart = CreateCountByWeekChart(lastMonth);
    const visitChart = CreateCountByVisitsChart(lastMonth);

    var blobs = new Array(4);
    var images = {};
    blobs[0] = cumulativeCountChart.getAs("image/png").setName("cumulativeCountChart");
    blobs[1] = cumulativeVisitChart.getAs("image/png").setName("cumulativeVisitChart");
    blobs[2] = countByWeekChart.getAs("image/png").setName("countByWeekChart");
    blobs[3] = visitChart.getAs("image/png").setName("visitChart");
    images["cumulativeCountChart"] = blobs[0];
    images["cumulativeVisitChart"] = blobs[1];
    images["countByWeekChart"] = blobs[2];
    images["visitChart"] = blobs[3];

    var body = "<h1>Woodside Monthly Report</h1>";
    body += "<h2>Attendance Year To Date</h2>";
    body += "<img src='cid:cumulativeCountChart' />";

    body += "<h2>Visits This Year</h2>";
    body += "<img src='cid:cumulativeVisitChart' />";

    body += "<h2>Attendance Month To Date</h2>";
    body += "<img src='cid:countByWeekChart' />";

    body += "<h2>Visits This Month</h2>";
    body += "<p>As in '24 dancers attended twice this month'</p>";
    body += `<img src="cid:visitChart" />`;

    body = AddWarnings(lastMonth, body);

    MailApp.sendEmail({
        to: "woodsideswingspokane@gmail.com",
        subject: `Monthly Report (${MONTHS[new Date().getMonth()]})`,
        htmlBody: body,
        inlineImages: images
    });
}

const CreateCumulativeCountChart = (stats: YearStats): GoogleAppsScript.Charts.Chart => {
    var data = Charts.newDataTable();
    data.addColumn(Charts.ColumnType.STRING, "Week");
    data.addColumn(Charts.ColumnType.NUMBER, "Attendance");
    let cumulative = 0;
    for(var i = 0; i < stats.countByWeek.length; i++) {
        cumulative += stats.countByWeek[i];
        data.addRow([`${i + 1}`, cumulative]);
    }

    return Charts.newLineChart()
        .setTitle("Cumulative Count")
        .setXAxisTitle("Week")
        .setYAxisTitle("Attendance")
        .setDimensions(800, 600)
        .setDataTable(data)
        .build();
}

const CreateCumulativeVisitChart = (stats: YearStats): GoogleAppsScript.Charts.Chart => {
    var data = Charts.newDataTable();
    data.addColumn(Charts.ColumnType.STRING, "Weeks");
    data.addColumn(Charts.ColumnType.NUMBER, "Dancers");
    for(var i = 1; i < stats.countByVisit.length; i++) {
        data.addRow([`${i}`, stats.countByVisit[i]]);
    }
    return Charts.newColumnChart()
        .setTitle("Visits This Year")
        .setXAxisTitle("Weeks Attended")
        .setYAxisTitle("Number of Dancers")
        .setDimensions(800, 600)
        .setDataTable(data)
        .build();
}

const CreateCountByWeekChart = (stats: MonthStats): GoogleAppsScript.Charts.Chart => {
    var countByWeekData = Charts.newDataTable();

    countByWeekData.addColumn(Charts.ColumnType.STRING, "Week");
    countByWeekData.addColumn(Charts.ColumnType.NUMBER, "Attendance");
    for (var i = 0; i < Object.keys(stats.countByWeek).length; i++) {
        countByWeekData.addRow([`Week ${i + 1}`, stats.countByWeek[i]]);
    }
    return Charts.newColumnChart()
        .setTitle("Count by week")
        .setXAxisTitle("Week")
        .setYAxisTitle("Attendance")
        .setDimensions(800, 600)
        .setDataTable(countByWeekData)
        .build();
}

const CreateCountByVisitsChart = (stats: MonthStats): GoogleAppsScript.Charts.Chart => {
    var peopleData = Charts.newDataTable();
    peopleData.addColumn(Charts.ColumnType.STRING, "Weeks");
    peopleData.addColumn(Charts.ColumnType.NUMBER, "Dancers");
    for(var i = 1; i < Object.keys(stats.countByVisit).length; i++) {
        peopleData.addRow([`${i}`, stats.countByVisit[i]]);
    }

    return Charts.newColumnChart()
        .setTitle("Visits This Month")
        .setXAxisTitle("Weeks Attended")
        .setYAxisTitle("Number of Dancers")
        .setDimensions(800, 600)
        .setDataTable(peopleData)
        .build();

}

const AddWarnings = (stats: MonthStats, body: string) : string => {
    if(stats.missingWaivers.length > 0 || stats.incompleteNames > 0 || stats.duplicateNames.length > 0) {
        body += "<h2>Warnings</h2>";
        if(stats.missingWaivers.length > 0) {
            body += "<h3>Missing Waivers</h3>";
            body += "<p>Dancers who attended, but do not have a signed waiver</p>";
            body += "<ul>";
            stats.missingWaivers.forEach(name => {
                body += `<li>${name}</li>`;
            });
            body += "</ul>";
        }
        if(stats.duplicateNames.length > 0) {
            body += `<h3>Duplicate Names</h3>`;
            body += "<ul>";
            stats.duplicateNames.forEach(name => {
                body += `<li>${name}</li>`;
            });
            body += "</ul>";
        }
        if(stats.incompleteNames > 0) {
            body += `<h3>Incomplete Names</h3>`;
            body += `<p>${stats.incompleteNames} people have incomplete names</p>`;
        }

    }
    return body;
}

const CountToWord = (count: number): string => {
    switch(count) {
        case 0: return "never";
        case 1: return "once";
        case 2: return "twice";
        case 3: return "three times";
        case 4: return "four times";
        default: return "A bunch";
    }
}