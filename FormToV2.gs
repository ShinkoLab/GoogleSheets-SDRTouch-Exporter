function FormToV2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    for (var i = 0; i < ss.getSheets().length; i++) {
        unFrozen(ss.getSheets()[i]);
        moveColumn(ss.getSheets()[i], "C1", "2");
        addColumn(ss.getSheets()[i], "3", "1")
        setText(ss.getSheets()[i], "D2", "Filter", 'bold')
        setFrozen(ss.getSheets()[i], "2", "5");
    }
}

function unFrozen(sheet) {
    //var sheet = ss.getSheets()[sheetno];
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
}

function moveColumn(sheet, src, dest) {
    //var sheet = ss.getSheets()[sheetno];
    var columnSpec = sheet.getRange(src);
    sheet.moveColumns(columnSpec, dest);
}

function addColumn(sheet, afterPosition, howMany) {
    //var sheet = ss.getSheets()[sheetno];
    sheet.insertColumnsAfter(afterPosition, howMany);
}

function setText(sheet, range, text, weight) {
    //var sheet = ss.getSheets()[sheetno];
    sheet.getRange(range).setValue(text);
    sheet.getRange(range).setFontWeight(weight);
}

function setFrozen(sheet, rows, columns) {
    //var sheet = ss.getSheets()[sheetno];
    sheet.setFrozenRows(rows);
    sheet.setFrozenColumns(columns);
}
