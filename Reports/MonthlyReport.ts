// import { TryGetSingleSheet } from "../Utils/SheetUtils.ts"
// import { GetAttendenceFile, SheetDetails } from "../Utils/WoodsideUtils.ts"
import { MONTHS } from "../Utils/Constants.ts"

export default function SendMonthlyReport() {
    const file = GetAttendenceFile();
    var monthStats: MonthStats[] = [];
    var weekNames: string[][] = [];
    for (var i = 0; i < MONTHS.length; i++) {
        var sheet = TryGetSingleSheet(file, MONTHS[i]);
        if(sheet === undefined) continue;

        const sheetDetails = new SheetDetails(sheet);

        const sheetSummary = SummarizeSheet(sheet, sheetDetails);
        const monthStat = SummarizeMonth(sheetSummary);
        monthStats[i] = monthStat; // keep month lined up with monthStats
        weekNames[i] = GetWeekNames(sheetDetails);
    }
    const yearStats = SummarizeYear(monthStats);
    SendEmail(monthStats, yearStats, weekNames);
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

const SummarizeSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet, sheetDetails: SheetDetails): RowSummary[] => {
    var result: RowSummary[] = [];
    var data = sheet.getDataRange().getValues();
    data.forEach((row, index) => {
        if (index === 0) { //skip header
            return;
        }
        var rowSummary: RowSummary = {
            hasWaiver: row[sheetDetails.WaiverColumn] !== "",
            fullName: `${row[sheetDetails.FirstNameColumn]} ${row[sheetDetails.LastNameColumn]}`.trim(),
            hasFullName: row[sheetDetails.FirstNameColumn] !== "" && row[sheetDetails.LastNameColumn] !== "",
            attended: []
        };
        for (var col = sheetDetails.FirstWeekColumn; col < row.length; col++) {
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

const SendEmail = (monthStats: MonthStats[], yearStats: YearStats, weekNames: string[][]) => {
    let lastMonth: MonthStats | undefined = undefined;
    let lastMonthWeeks: string[] = [];

    for(let i = monthStats.length - 1; i >= 0; i--) {
        if(monthStats[i] !== undefined && monthStats[i].countByWeek[0] > 0) {
            lastMonth = monthStats[i];
            lastMonthWeeks = weekNames[i];
            break;
        }
    }
    if(lastMonth === undefined) {
        Logger.log("Nothing to send");
        return;
    }

    const cumulativeCountChart = CreateCumulativeCountChart(yearStats);
    const cumulativeVisitChart = CreateCumulativeVisitChart(yearStats);
    const countByWeekChart = CreateCountByWeekChart(lastMonth, lastMonthWeeks);
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

    body += "<h2>Visits Last Month</h2>";
    body += "<p>As in '24 dancers attended twice last month'</p>";
    body += `<img src="cid:visitChart" />`;

    body = AddWarnings(lastMonth, body);

    MailApp.sendEmail({
        to: "woodsideswingspokane@gmail.com",
        cc: "dstbstr17@gmail.com",
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

const CreateCountByWeekChart = (stats: MonthStats, names: string[]): GoogleAppsScript.Charts.Chart => {
    var countByWeekData = Charts.newDataTable();

    countByWeekData.addColumn(Charts.ColumnType.STRING, "Week");
    countByWeekData.addColumn(Charts.ColumnType.NUMBER, "Attendance");
    for (var i = 0; i < Object.keys(stats.countByWeek).length; i++) {
        countByWeekData.addRow([names[i], stats.countByWeek[i]]);
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

const GetWeekNames = (sheetDetails: SheetDetails): string[] => {
    var result: string[] = [];
    for(var i = sheetDetails.FirstWeekColumn; i < Object.keys(sheetDetails.Lut).length; i++) {
        const val = Object.keys(sheetDetails.Lut)[i];
        var date = new Date(val);
        if(isNaN(date.getTime())) {
            throw new Error(`Could not parse date ${val}`);
        }

        //add "Thu 3" or "Mon 7" to the result
        const options: Intl.DateTimeFormatOptions = {weekday: "short", day: "numeric"};
        result.push(date.toLocaleDateString("en-US", options));
    }
    return result;
}