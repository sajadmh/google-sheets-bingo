function GetSheetName() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function onEdit(e) {

    const notation = e.range.getA1Notation();

    if (notation === "D15") { // Centers "free space" for all cards if checkbox is checked
        const sheet = e.range.getSheet();
        if (sheet.getName() === "Settings") {
            const value = e.range.getValue();
            if (value === true) {
                const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
                for (let i = 0; i < sheets.length; i++) {
                    const sh = sheets[i];
                    if (sh.getName() !== "Settings") {
                        const freeSpace = centerFreeSpace();
                        sh.getRange("D9").setValue(freeSpace); // Range of each board where values are assigned
                    }
                }
            }
        }
    }

    if (notation === "D4") { // Checkbox that generates random assignment
        const sheet = e.range.getSheet();
        if (sheet.getName() === "Settings") {
            const value = e.range.getValue();
            if (value === true) {
                const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
                for (let i = 0; i < sheets.length; i++) {
                    const sh = sheets[i];
                    if (sh.getName() !== "Settings") {
                        const randomValues = getRandomValues();
                        sh.getRange("B7:F11").setValues(randomValues); // Range of each board where values are assigned
                    }
                }
            }
        }
    }

    if (notation === "A1") { // Checkbox for individual cards
        const value = e.range.getValue();
        if (value === true) {
            const sheet = e.range.getSheet();
            const randomValues = getRandomValues();
            sheet.getRange("B7:F11").setValues(randomValues);
        }
    }

    /* Modified snippet from Daniel Kilcoyne https://github.com/dpkilcoyne/google-sheets-bingo */

    const ss = e.source;
    const value = e.value;
    const sheet = ss.getActiveSheet();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();

    if (row >= 2 & row <= 6 & col >= 8 & col <= 12) {
        if (value <= 99) {
            sheet.getRange(row + 5, col - 6).setBackground('#96c265').setFontColor('#FFFFFF'); // Sets the background color for when you fill in a square with the mini card
        } else if (value >= 999) {
            sheet.getRange(row + 5, col - 6).setBackground('#FFFFFF').setFontColor('#000000');
        }
    }
}

function getRandomValues() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    const values = sheet.getRange("B4:B86").getValues().filter(item => item[0] !== ""); // Range of values in Settings to be assigned
    values.sort((prev, cur) => Math.random() - .5);
    const randomValuesArray = [];
    for (let i = 0; i < 5; i++) {
        randomValuesArray[i] = values.slice(i * 5, (i + 1) * 5);
    }
    return randomValuesArray;
}

function centerFreeSpace() {
    const freeSpace = "FREE ★ SPACE";
    return freeSpace;
}
