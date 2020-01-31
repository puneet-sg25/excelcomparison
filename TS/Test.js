"use strict";
exports.__esModule = true;
var Excel = require("exceljs");
var XLSX = require("xlsx");
var wb = new Excel.Workbook();
var workbook = XLSX.readFile("./sample.xlsx");
var sheets = workbook.Sheets;
var sheetNames = workbook.SheetNames;
for (var i = 0; i < workbook.SheetNames.length; ++i) {
    var sheet = workbook.Sheets[workbook.SheetNames[i]];
    console.log(workbook.SheetNames[i]);
}
wb.xlsx.readFile("./sample.xlsx").then(function () {
    var sh = wb.getWorksheet("Sheet1");
    var sh1 = wb.getWorksheet("Sheet2");
    for (var i_1 = 1; i_1 <= sh.rowCount; i_1++) {
        console.log(sh.getRow(i_1).getCell(1).value);
        console.log(sh.getRow(i_1).getCell(2).value);
    }
    for (var i_2 = 1; i_2 <= sh1.rowCount; i_2++) {
        console.log(sh1.getRow(i_2).getCell(1).value);
        console.log(sh1.getRow(i_2).getCell(2).value);
    }
});
