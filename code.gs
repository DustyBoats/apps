function main() {
    checkAndUpdateSpreadsheets();
    updateOriginalSpreadsheetsBasedOnMasterSheet();

    var sourceSpreadsheetId = '1PW_XWebquFPTJnS7eM5Fz9IgAVXyeMqUB7OxrPJCzLU';
    var targetSheetName = 'Master';
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    var targetSheet = sourceSpreadsheet.getSheetByName(targetSheetName);

    checkMasterSheetAndAddTimestamps(targetSheet);
    createRecruiterDashboard();
}

function checkAndUpdateSpreadsheets() {
    var sourceSpreadsheetId = '1PW_XWebquFPTJnS7eM5Fz9IgAVXyeMqUB7OxrPJCzLU';
    var sourceSheetName = 'PERM Recruitment File ID\'s';
    var targetSheetName = 'Master';
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
    var targetSheet = sourceSpreadsheet.getSheetByName(targetSheetName);
    var sourceData = sourceSheet.getDataRange().getValues();
    var targetData = [];
    if (targetSheet.getLastRow() > 1) { // Check if there's more than one row
        targetData = targetSheet.getRange(2, 25, targetSheet.getLastRow() - 1).getValues(); // Column Y data for checking
    }

    Logger.log('Starting checkAndUpdateSpreadsheets function...');

    for (var i = 1; i < sourceData.length; i++) {
        var id = sourceData[i][17]; // assuming IDs are in column R
        Logger.log('Processing spreadsheet ID: ' + id);
        var spreadsheet = SpreadsheetApp.openById(id);
        var sheet = spreadsheet.getSheets()[0]; // assuming you need the first sheet in each spreadsheet
        var data = sheet.getDataRange().getValues();

        for (var j = 0; j < data.length; j++) {
            var state = data[j][0].toLowerCase();
            var timestamp = new Date();

            var rowId = id + "_" + j; // Create a Row ID using File ID and Row Number

            // Check if Row ID already exists in the Master sheet (Column Y)
            var found = false;
            for (var l = 0; l < targetData.length; l++) {
                if (targetData[l][0] == rowId) {
                    found = true;
                    break;
                }
            }

            // If Row ID already exists or state does not contain 'Phone Screen Required', skip appending
            if (found || state != 'requires phone screen') {
                Logger.log('Row ' + rowId + ' already exists in the Master sheet or state does not equal Phone Screen Required.');
                continue;
            }

            Logger.log('Processing row in spreadsheet ID: ' + id);

            // Append the values from A-C to the Master sheet
            var newRow = data[j].slice(0, 3);
            newRow[24] = rowId; // Set the Row ID in column Y
            targetSheet.appendRow(newRow);

            // Copying values from columns A-H (indices 0-7) in the 'PERM Recruitment File ID's' sheet to columns H-N in the 'Master' sheet
            for (var k = 0; k < 8; k++) {
                targetSheet.getRange(targetSheet.getLastRow(), k+8).setValue(sourceData[i][k]);
            }
        }
    }

    Logger.log('Finished checkAndUpdateSpreadsheets function...');
}

function checkMasterSheetAndAddTimestamps(targetSheet) {
    var data = targetSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { // skip header row
        var state = data[i][0]; // assuming state is in column A
        var timestamp = new Date();
        checkStateAndAddTimestamp(state, targetSheet, i+1, timestamp);
    }
}

function checkStateAndAddTimestamp(masterState, targetSheet, rowIndex, timestamp) {
    // Convert masterState to lower case
    masterState = masterState.toLowerCase();

    // Check if the state in the Master sheet contains 'requires phone screen', 'pending' or 'complete' and add a timestamp in the corresponding column.
    if (masterState == 'requires phone screen') {
        // Check if column AA (column index 27) already contains a value
        var cell = targetSheet.getRange(rowIndex, 27);
        if (!cell.getValue()) {
            cell.setValue(timestamp);
            Logger.log('State is "requires phone screen". Added timestamp to column AA for this row.');
        }
    }
    if (masterState.includes('pending')) {
        // Check if column AB (column index 28) already contains a value
        var cell = targetSheet.getRange(rowIndex, 28);
        if (!cell.getValue()) {
            cell.setValue(timestamp);
            Logger.log('State contains "pending". Added timestamp to column AB for this row.');
        }
    } else if (masterState.includes('complete')) {
        // Check if column AC (column index 29) already contains a value
        var cell = targetSheet.getRange(rowIndex, 29);
        if (!cell.getValue()) {
            cell.setValue(timestamp);
            Logger.log('State contains "complete". Added timestamp to column AC for this row.');
        }
    } else {
        Logger.log('State does not contain "requires phone screen", "pending" or "complete". No timestamp added.');
    }

    Logger.log('Checked and updated timestamps for row in the Master sheet.');
}

function toTitleCase(str) {
    return str.replace(
        /\w\S*/g,
        function(txt) {
            return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
        }
    );
}

function updateOriginalSpreadsheetsBasedOnMasterSheet() {
    Logger.log('Starting updateOriginalSpreadsheetsBasedOnMasterSheet function...');

    var sourceSpreadsheetId = '1PW_XWebquFPTJnS7eM5Fz9IgAVXyeMqUB7OxrPJCzLU';
    var targetSheetName = 'Master';
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    var targetSheet = sourceSpreadsheet.getSheetByName(targetSheetName);

    var targetData = targetSheet.getDataRange().getValues();
    Logger.log('Loaded data from Master sheet, total rows: ' + targetData.length);

    for (var i = 1; i < targetData.length; i++) { // Starts from the second row
        var status = targetData[i][0].toLowerCase();
        var rowId = targetData[i][24]; // Row ID is in column Y
        if (!(status.includes('complete') || status.includes('pending'))) {
            Logger.log('Skipping row ' + (i+1) + ', status does not contain complete or pending.');
            continue; // Skip if status does not contain 'complete' or 'pending'
        }

        Logger.log('Processing row ' + (i+1) + ', status: ' + status);
        
        var rowIdParts = rowId.split("_");
        var originalRowNumber = parseInt(rowIdParts.pop()); // Row number is the last part
        var originalSpreadsheetId = rowIdParts.join("_"); // The remaining parts form the spreadsheet ID

        Logger.log('Opening original spreadsheet with ID: ' + originalSpreadsheetId);

        var originalSpreadsheet = SpreadsheetApp.openById(originalSpreadsheetId);
        var originalSheet = originalSpreadsheet.getSheets()[0]; // assuming the first sheet is to be updated

        var originalData = originalSheet.getDataRange().getValues();
        Logger.log('Loaded data from original sheet, total rows: ' + originalData.length);

        // Update the status in the original sheet
        var titleCaseStatus = toTitleCase(status);
        Logger.log('Updating status in original sheet row ' + (originalRowNumber+1) + ' to ' + titleCaseStatus);
        originalSheet.getRange(originalRowNumber + 1, 1).setValue(titleCaseStatus);

        // Append values from D-G in Master sheet to column E in original sheet, if they aren't empty
var additionalData = [];
for (var j = 3; j < 7; j++) { // Indices 3-6 for columns D-G
    if (targetData[i][j] != '') {
        additionalData.push(targetData[i][j]);
    }
}

if (additionalData.length > 0) {
    var originalDataInColumnE = originalData[originalRowNumber][4].split(' ');

    additionalData.forEach(function(item) {
        if (!originalDataInColumnE.includes(item)) {
            originalDataInColumnE.push(item);
        }
    });

    var updatedDataInColumnE = originalDataInColumnE.join(' ');

    if (updatedDataInColumnE != originalData[originalRowNumber][4]) {
        Logger.log('Appending additional data to column E in original sheet row ' + (originalRowNumber+1));
        originalData[originalRowNumber][4] = updatedDataInColumnE; // Column E is index 4
        originalSheet.getRange(originalRowNumber + 1, 5).setValue(updatedDataInColumnE);
    }
}
    }

    Logger.log('Finished updateOriginalSpreadsheetsBasedOnMasterSheet function...');
}




function createRecruiterDashboard() {
    console.log('Function createRecruiterDashboard started...');

    var spreadsheetId = '1PW_XWebquFPTJnS7eM5Fz9IgAVXyeMqUB7OxrPJCzLU';
    var sheetName = 'Master';
    var dashboardName = 'Recruiter Dashboard';

    console.log(`Connecting to spreadsheet with ID ${spreadsheetId}...`);
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    console.log('Getting data sheet...');
    var sheet = spreadsheet.getSheetByName(sheetName);

    console.log('Getting dashboard sheet...');
    var dashboard = spreadsheet.getSheetByName(dashboardName);

    if (!dashboard) {
        console.log('Dashboard sheet does not exist. Creating...');
        dashboard = spreadsheet.insertSheet(dashboardName);
    }

    console.log('Removing existing charts...');
    var existingCharts = dashboard.getCharts();
    for (var i = 0; i < existingCharts.length; i++) {
        dashboard.removeChart(existingCharts[i]);
    }

    console.log('Fetching sheet data...');
    var data = sheet.getDataRange().getValues();

    var recruiters = {};

    var now = new Date();
    var currentWeekStart = new Date(now);
    currentWeekStart.setDate(now.getDate() - now.getDay());
    currentWeekStart.setHours(0, 0, 0, 0); 
    var currentMonth = now.getMonth();

    console.log('Iterating over rows in Master sheet...');
    for (var i = 1; i < data.length; i++) {
        var nameAndState = data[i][0]; 
        var timestamp = new Date(data[i][26]); 
        var pendingTimestamp = new Date(data[i][27]);
        var completeTimestamp = new Date(data[i][28]);

var recruiter = toTitleCase(nameAndState.split(' ')[0]);

       // var recruiter = nameAndState.split(' ')[0].toLowerCase();
        var state = nameAndState.split(' ')[1].toLowerCase();

if (recruiter === 'requires' || recruiter === 'recruiter' || recruiter === 'Recruiter' || recruiter === 'recruiter' || recruiter === 'Requires') {
    continue;  // Skip this row if recruiter is "requires" or "recruiter" or "Recruiter"
}

        if (!recruiters[recruiter]) {
            recruiters[recruiter] = {
               // scheduled: { week: 0, month: 0 },
                completed: { week: 0, month: 0 },
                pending: { week: 0, month: 0 },  // new
                duration: 0,
                count: 0
            };
        }

           Logger.log(`Processing row ${i}: recruiter = ${recruiter}, state = ${state}, pendingTimestamp = ${pendingTimestamp}, completeTimestamp = ${completeTimestamp}`);

        pendingTimestamp.setHours(0, 0, 0, 0);
        completeTimestamp.setHours(0, 0, 0, 0);

        if (state.toLowerCase().includes('pending') || isNaN(completeTimestamp.getTime())) { 
            if (pendingTimestamp >= currentWeekStart) recruiters[recruiter].pending.week += 1; 
            if (pendingTimestamp.getMonth() == currentMonth) recruiters[recruiter].pending.month += 1; 
        } else if (state.toLowerCase().includes('complete')) {
            if (completeTimestamp >= currentWeekStart) recruiters[recruiter].completed.week += 1;
            if (completeTimestamp.getMonth() == currentMonth) recruiters[recruiter].completed.month += 1;
            recruiters[recruiter].duration += completeTimestamp - pendingTimestamp;
            recruiters[recruiter].count += 1; 
        }
    }

    console.log('Clearing Dashboard sheet...');
    dashboard.clear();

    console.log('Creating headers in Dashboard sheet...');
    var headers = ['Recruiter', 'Scheduled This Week', 'Completed This Week', 'Scheduled This Month', 'Completed This Month', 'Average Time to Complete'];
    dashboard.appendRow(headers);

    console.log('Iterating through each recruiter...');
    Object.keys(recruiters).forEach(function (recruiter) {
       // var scheduledWeek = recruiters[recruiter].scheduled.week;
        //var scheduledMonth = recruiters[recruiter].scheduled.month;
        var completedWeek = recruiters[recruiter].completed.week;
        var completedMonth = recruiters[recruiter].completed.month;
        var pendingWeek = recruiters[recruiter].pending.week;  // new
        var pendingMonth = recruiters[recruiter].pending.month;  // new
        var totalDuration = recruiters[recruiter].duration;
        var completeCount = recruiters[recruiter].count;

        console.log(`Processing recruiter ${recruiter}: pendingWeek = ${pendingWeek}, pendingMonth = ${pendingMonth}, completedWeek = ${completedWeek}, completedMonth = ${completedMonth}, totalDuration = ${totalDuration}`);

        var avgDuration = completeCount > 0 ? totalDuration / (completeCount * 24 * 60 * 60 * 1000) : 0;

        var row = [recruiter, pendingWeek, completedWeek, pendingMonth, completedMonth, avgDuration];
        dashboard.appendRow(row);
    });

    console.log('Creating bar chart...');
    var lastRow = dashboard.getLastRow();
    var lastColumn = dashboard.getLastColumn();
    var chartRange = dashboard.getRange(2, 1, lastRow - 1, lastColumn);
    //var chartRange = dashboard.getRange(1, 1, lastRow, lastColumn);

 var comboChart = dashboard.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(chartRange)
    .setPosition(lastRow + 1, 1, 5, 0)
    .setOption('title', 'Recruiter Dashboard')
    .setOption('legend', { position: 'top' })
    .setOption('colors', ['#4285F4', '#FF5722'])
    .setOption('animation.duration', 1000)
    .setOption('seriesType', 'bars')
    .setOption('annotations', { alwaysOutside: false })
    .setOption('annotations.textStyle.color', '#000')
    .setOption('annotations.textStyle.fontSize', 14)
    .setOption('series', {
        0: { labelInLegend: 'Pending This Week', type: 'bars', dataLabel: 'value' },
        //1: { labelInLegend: 'Scheduled This Week', type: 'line', dataLabel: 'value' },
        1: { labelInLegend: 'Completed This Week', type: 'bars', dataLabel: 'value'  },
        2: { labelInLegend: 'Pending This Month', type: 'bars', dataLabel: 'value' },
        //4: { labelInLegend: 'Scheduled This Month', type: 'bars' },
        3: { labelInLegend: 'Completed This Month', type: 'bars', dataLabel: 'value' },
        4: { labelInLegend: 'Average Time to Complete', type: 'line', dataLabel: 'none' }
    })
    .setOption('axes', {
        x: {
            0: { side: 'top', label: 'Recruiter' }
        },
        y: {
            0: { side: 'left', label: 'Count' },
            1: { side: 'right', label: 'Days' }
        }
    })
    .setOption('width', 1000)
    .setOption('height', 400)
    .setOption('vAxis', { format: 'short', gridlines: { count: 4 }, minorGridlines: { count: 1 } }) // Added gridlines and minorGridlines for the vertical axis
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45, gridlines: { count: -1 }, minorGridlines: { count: 1 } }) // Added gridlines and minorGridlines for the horizontal axis
    .build();

dashboard.insertChart(comboChart);
console.log('Function createRecruiterDashboard finished...');
}

// Function to format a Date object as a string in the format 'yyyy-MM-dd'
function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) 
        month = '0' + month;
    if (day.length < 2) 
        day = '0' + day;

    return [year, month, day].join('-');
}

// Function to calculate the average number of days between dates in a list, for the past numDays days
function calculateAverageDays(dates, numDays) {
    if (dates.length < 2) {
        return '-';
    }

    // Sort the dates in ascending order
    dates.sort();

    // Convert the dates to milliseconds since the Unix epoch
    var timestamps = dates.map(function(date) {
        return new Date(date).getTime();
    });

    // Get the timestamp numDays days ago
    var cutoffTimestamp = new Date().getTime() - numDays * 24 * 60 * 60 * 1000;

    // Filter the timestamps to only include those within the past numDays days
    var recentTimestamps = timestamps.filter(function(timestamp) {
        return timestamp >= cutoffTimestamp;
    });

    if (recentTimestamps.length < 2) {
        return '-';
    }

    // Calculate the differences between consecutive timestamps, then take the average
    var diffs = [];
    for (var i = 1; i < recentTimestamps.length; i++) {
        diffs.push(recentTimestamps[i] - recentTimestamps[i - 1]);
    }
    var avgDiff = diffs.reduce(function(a, b) {
        return a + b;
    }) / diffs.length;

    // Convert the average difference from milliseconds to days and return it
    return avgDiff / (24 * 60 * 60 * 1000);
}
