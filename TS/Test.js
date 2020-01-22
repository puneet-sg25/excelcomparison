"use strict";
exports.__esModule = true;
var Excel = require("exceljs");
var wb = new Excel.Workbook();
wb.xlsx.readFile("./sample.xlsx").then(function () {
    var sh = wb.getWorksheet("Sheet1");
    //sh.getRow(3).getCell(2).value = 32;
    //wb.xlsx.writeFile("./sample.xlsx");
    //console.log("Row-3 | Cell-2 - "+sh.getRow(3).getCell(2).value);
    console.log(sh.rowCount);
    //Get all the rows data [1st and 2nd column]
    for (var i = 1; i <= sh.rowCount; i++) {
        console.log(sh.getRow(i).getCell(1).value);
        console.log(sh.getRow(i).getCell(2).value);
    }
});
