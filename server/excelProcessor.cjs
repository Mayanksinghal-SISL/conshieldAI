const XLSX = require('xlsx');

// Helper function for safe numeric conversion - Keep this here
const getNumericValue = (text) => {
    const raw = String(text).trim().replace(/[$,%]/g, '');
    return (!isNaN(raw) && raw !== '' && raw.toUpperCase() !== 'NA') ? Number(raw) : null;
};

/**
 * Processes the raw Excel file ArrayBuffer to extract data,
 * perform calculations, and generate comments.
 * This function is designed to run on the server-side.
 * It returns the processed data, not HTML elements.
 *
 * @param {ArrayBuffer} arrayBuffer The raw data from the Excel file.
 * @returns {Object} An object containing client data for display and combined date headers.
 */
function processExcelDataServer(arrayBuffer) {
    const data = new Uint8Array(arrayBuffer);
    let workbook;
    try {
        workbook = XLSX.read(data, { type: 'array' });
    } catch (e) {
        console.error("Server: Error reading Excel file:", e);
        throw new Error("Could not read the Excel file. Please ensure it's a valid .xlsx file.");
    }

    // --- FIRST LOG: Shows ALL sheet names found in the workbook ---
    console.log("Server: Actual Sheet Names from Workbook:", workbook.SheetNames);

    // --- DECLARE sheetNames HERE (THIS LINE IS CRUCIAL) ---
const sheetNames = workbook.SheetNames.filter(name => /^(Sheet1|[A-Za-z]{3,}-\d{2})$/.test(name));

    // --- SECOND LOG: Shows sheet names AFTER filtering by your pattern ---
    // This log MUST come AFTER the 'const sheetNames =' line
    console.log("Server: Filtered Sheet Names (expected pattern):", sheetNames);


    if (sheetNames.length === 0) {
        console.warn("Server: No sheets found matching the expected pattern (e.g., 'Jan-23').");
        throw new Error("No relevant data sheets found in the Excel file. Please check sheet names.");
    }

    const clientSet = new Set();
    const dayColsBySheet = {};
    const clientRowMapBySheet = {};

    const today = new Date();
    const yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0);

    const fromDate = new Date(yesterday);
    fromDate.setDate(fromDate.getDate() - 29); // Last 50 days
    fromDate.setHours(0, 0, 0, 0);

    const monthIndexMap = {
        jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
        jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
    };

    sheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet || !sheet['!ref']) {
            console.warn(`Server: Sheet "${sheetName}" is empty or invalid.`);
            return;
        }
        const range = XLSX.utils.decode_range(sheet['!ref']);
        const sheetYear = parseInt('20' + sheetName.slice(-2));
        const dayCols = [];
        const rowMap = {};

        for (let col = 6; col <= range.e.c; col++) {
            const cell = sheet[XLSX.utils.encode_cell({ c: col, r: 0 })];
            if (!cell || !cell.v) continue;

            let cellDate = null;

            if (cell.t === 'n') {
                const parsed = XLSX.SSF.parse_date_code(cell.v);
                if (parsed) cellDate = new Date(parsed.y, parsed.m - 1, parsed.d);
            } else if (cell.t === 's') {
                const match = cell.v.match(/(\d{1,2})-([A-Za-z]{3})(?:-(\d{2}))?/);
                if (match) {
                    const [_, d, mmm, yy] = match;
                    const monthNum = monthIndexMap[mmm.toLowerCase()];
                    const year = yy ? parseInt('20' + yy) : sheetYear;
                    cellDate = new Date(year, monthNum, parseInt(d));
                }
            }

            if (cellDate) {
                const cellDateNorm = new Date(cellDate); cellDateNorm.setHours(0,0,0,0);
                const fromDateNorm = new Date(fromDate); fromDateNorm.setHours(0,0,0,0);
                const yesterdayNorm = new Date(yesterday); yesterdayNorm.setHours(0,0,0,0);

                if(cellDateNorm >= fromDateNorm && cellDateNorm <= yesterdayNorm) {
                    dayCols.push({ col, date: cellDate });
                }
            }
        }
        dayColsBySheet[sheetName] = dayCols;

        for (let row = 1; row <= range.e.r; row++) {
            const cell = sheet[`B${row + 1}`];
            if (cell && cell.v) {
                const client = String(cell.v).trim();
                if (client) {
                    clientSet.add(client);
                    rowMap[client] = row;
                }
            }
        }
        clientRowMapBySheet[sheetName] = rowMap;
    });

    const combinedDateCols = [];
    sheetNames.forEach(sheetName => {
        const sheetDayCols = dayColsBySheet[sheetName] || [];
        sheetDayCols.forEach(dayCol => {
            const isDuplicate = combinedDateCols.some(existing =>
                existing.date.getTime() === dayCol.date.getTime()
            );
            if (!isDuplicate) {
                combinedDateCols.push(dayCol);
            }
        });
    });
    combinedDateCols.sort((a, b) => a.date.getTime() - b.date.getTime());
    const combinedDateHeaders = combinedDateCols.map(dayCol =>
        dayCol.date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit' })
    );

    const clientDataForDisplay = [];
    let hasNonZeroMonthlyCount = false;
    let hasNonZeroFortnightlyCount = false;
    let hasNonZeroWeeklyCount = false;

    Array.from(clientSet).sort().forEach(client => {
        const clientDailyValues = [];
        combinedDateCols.forEach(({ col, date }) => {
            let value = '';
            for (const sheetName of sheetNames) {
                const sheet = workbook.Sheets[sheetName];
                const rowIndex = clientRowMapBySheet[sheetName]?.[client];
                if (typeof rowIndex !== 'undefined') {
                    const cell = sheet[XLSX.utils.encode_cell({ c: col, r: rowIndex })];
                    let cellDateInSheet = null;
                    const headerCell = sheet[XLSX.utils.encode_cell({ c: col, r: 0 })];
                    if (headerCell && headerCell.v) {
                        if (headerCell.t === 'n') {
                            const parsed = XLSX.SSF.parse_date_code(headerCell.v);
                            if (parsed) cellDateInSheet = new Date(parsed.y, parsed.m - 1, parsed.d);
                        } else if (headerCell.t === 's') {
                            const match = headerCell.v.match(/(\d{1,2})-([A-Za-z]{3})(?:-(\d{2}))?/);
                            if (match) {
                                const [_, d, mmm, yy] = match;
                                const monthNum = monthIndexMap[mmm.toLowerCase()];
                                const year = yy ? parseInt('20' + yy) : new Date().getFullYear();
                                cellDateInSheet = new Date(year, monthNum, parseInt(d));
                            }
                        }
                    }
                    const normalizedCellDateInSheet = cellDateInSheet ? new Date(cellDateInSheet).setHours(0,0,0,0) : null;
                    const normalizedDate = new Date(date).setHours(0,0,0,0);

                    if (normalizedCellDateInSheet !== null && normalizedCellDateInSheet === normalizedDate) {
                        value = (cell && cell.v !== undefined) ? cell.v : '';
                        break;
                    }
                }
            }
            clientDailyValues.push(value);
        });

        const validMonthlyValues = clientDailyValues.slice(-30).filter(val => {
            const cleaned = (typeof val === 'string') ? val.replace(/[$,]/g, '').trim() : val;
            return !isNaN(cleaned) && cleaned !== '' && cleaned !== null && String(cleaned).toUpperCase() !== 'NA';
        });
        const monthlyCount = validMonthlyValues.length;
        if (monthlyCount > 0) hasNonZeroMonthlyCount = true;

        const last15Numeric = clientDailyValues.slice(-15).filter(val => {
            const raw = String(val).trim().replace(/[$,]/g, '');
            return (!isNaN(raw) && raw !== '' && raw.toUpperCase() !== 'NA');
        });
        const fortnightlyCount = last15Numeric.length;
        if (fortnightlyCount > 0) hasNonZeroFortnightlyCount = true;

        const last7Numeric = clientDailyValues.slice(-7).filter(val => {
            const raw = String(val).trim().replace(/[$,]/g, '');
            return (!isNaN(raw) && raw !== '' && raw.toUpperCase() !== 'NA');
        });
        const weeklyCount = last7Numeric.length;
        if (weeklyCount > 0) hasNonZeroWeeklyCount = true;

        const cleanedLast30 = clientDailyValues.slice(-30).map(val => getNumericValue(val)).filter(v => v !== null);
        let monthlyAvgValue = 'N/A';
        if (cleanedLast30.length > 0) {
            const sum = cleanedLast30.reduce((a, b) => a + b, 0);
            monthlyAvgValue = (sum / cleanedLast30.length).toFixed(2);
        }

        const last15NumericValues = clientDailyValues.slice(-15).map(val => getNumericValue(val)).filter(v => v !== null);
        let fortnightlyAvgValue = 'N/A';
        if (last15NumericValues.length > 0) {
            const sum = last15NumericValues.reduce((a, b) => a + b, 0);
            fortnightlyAvgValue = (sum / last15NumericValues.length).toFixed(2);
        }

        const last7NumericValues = clientDailyValues.slice(-7).map(val => getNumericValue(val)).filter(v => v !== null);
        let weeklyAvgValue = 'N/A';
        if (last7NumericValues.length > 0) {
            const sum = last7NumericValues.reduce((a, b) => a + b, 0);
            weeklyAvgValue = (sum / last7NumericValues.length).toFixed(2);
        }

        const yesterdayDataValue = clientDailyValues[clientDailyValues.length - 1] || '';
        const yesterdayVal = getNumericValue(yesterdayDataValue);
        const diffPctValue = 5;

        const comments = [];
        let commentClass = '';

        if (yesterdayVal === null || yesterdayDataValue === '') {
            comments.push("Yesterday Data Blank");
            commentClass = 'comment-grey';
        } else {
            let hasDiff = false;
            const checkDifference = (avg, type) => {
                if (avg !== null && yesterdayVal !== null && diffPctValue !== null) {
                    const diff = Math.abs(avg - yesterdayVal);
                    if (avg !== 0 && (diff / avg) * 100 > diffPctValue) {
                        comments.push(`${type} difference more than ${diffPctValue}%`);
                        hasDiff = true;
                    }
                }
            };
            checkDifference(getNumericValue(monthlyAvgValue), 'Monthly');
            checkDifference(getNumericValue(fortnightlyAvgValue), 'Fortnightly');
            checkDifference(getNumericValue(weeklyAvgValue), 'Weekly');

            if (hasDiff) {
                commentClass = 'comment-red';
            } else {
                comments.push("No Action required");
                commentClass = 'comment-green';
            }
        }

        clientDataForDisplay.push({
            client: client,
            dailyValues: clientDailyValues,
            monthlyCount: monthlyCount,
            fortnightlyCount: fortnightlyCount,
            weeklyCount: weeklyCount,
            diffPct: `${diffPctValue}%`,
            monthlyAvg: monthlyAvgValue,
            fortnightlyAvg: fortnightlyAvgValue,
            weeklyAvg: weeklyAvgValue,
            yesterdayData: yesterdayDataValue,
            comments: comments.join(', '),
            commentClass: commentClass
        });
    });

    if (clientDataForDisplay.length === 0) {
        console.warn("Server: No client data found for display after processing.");
        throw new Error("No relevant client data found to display.");
    }

    return {
        clientDataForDisplay,
        combinedDateHeaders,
        hasNonZeroMonthlyCount, // Also return these flags for the frontend to use
        hasNonZeroFortnightlyCount,
        hasNonZeroWeeklyCount
    };
}

module.exports = {
    processExcelDataServer
};