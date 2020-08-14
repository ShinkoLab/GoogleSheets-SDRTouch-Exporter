function FormToV2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    for (var i = 0; i < ss.getSheets().length; i++) {
        unFrozen(ss.getSheets()[i]);
        moveColumn(ss.getSheets()[i], "C1", "2");
        addColumn(ss.getSheets()[i], "3", "1")
        setText(ss.getSheets()[i], "D2", "Filter", 'bold')
        setFrozen(ss.getSheets()[i], "2", "5");
        Logger.log("done: " + ss.getSheets()[i].getSheetName());
    }
    Logger.log("Conversion complete!")
}

function unFrozen(sheet) {
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
}

function moveColumn(sheet, src, dest) {
    var columnSpec = sheet.getRange(src);
    sheet.moveColumns(columnSpec, dest);
}

function addColumn(sheet, afterPosition, howMany) {
    sheet.insertColumnsAfter(afterPosition, howMany);
}

function setText(sheet, range, text, weight) {
    sheet.getRange(range).setValue(text);
    sheet.getRange(range).setFontWeight(weight);
}

function setFrozen(sheet, rows, columns) {
    sheet.setFrozenRows(rows);
    sheet.setFrozenColumns(columns);
}
