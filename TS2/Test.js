/*var excel_compare = require("excel-compare");
 
excel_compare({
    file1: 'sample.xlsx', // file1 is the main excel to compare with
    file2: 'sample_2.xlsx', // file2 is the file for compare
    column_file1: {
        column: [1],
        join: ''
    },
    column_file2: {
        column: [1,2],
        join: '-'
    }
})*/
var Excel = require('exceljs');
var wb1 = new Excel.Workbook();
var wb2 = new Excel.Workbook();
wb1.xlsx.readFile("./first.xlsx").then(function () {
    var sh1 = wb1.getWorksheet("Sheet1");
    console.log(sh1.rowCount);
    // for (let i = 1; i <= sh1.rowCount; i++) { 
    //      console.log(sh1.getRow(i).getCell(1).value);
    //     console.log(sh1.getRow(i).getCell(2).value);
    //  }
    wb2.xlsx.readFile("./second.xlsx").then(function () {
        var sh2 = wb2.getWorksheet("Sheet2");
        console.log(sh2.rowCount);
        if (compareTwoSheets(sh1, sh2)) {
            console.log("The Sheets are equal");
        }
        else {
            console.log("The Sheets are not equal");
        }
        function compareTwoSheets(sh1, sh2) {
            var firstRow1 = sh1.getRow(1);
            var lastRow1 = sh1.getRow(sh1.rowCount);
            var equalSheets = true;
            for (var i = firstRow1; i <= lastRow1; i++) {
                var row1 = sh1.getRow(i);
                var row2 = sh2.getRow(i);
                if (!compareTwoRows(row1, row2)) {
                    equalSheets = false;
                    console.log("Not equal rows");
                    break;
                }
                else {
                    console.log("Equal rows");
                }
            }
            return equalSheets;
        }
        function compareTwoRows(row1, row2) {
            if ((row1 == null) && (row2 == null)) {
                return true;
            }
            else if ((row1 == null) || (row2 == null)) {
                return false;
            }
            //var firstCell1: number = row1.getFirstCellNum();
            //var lastCell1: number = row1.getLastCellNum();
            var equalRows = true;
            // Compare all cells in a row
            for (var i = 1; i <= sh1.cellCount; i++) {
                var cell1 = row1.getCell(i);
                var cell2 = row2.getCell(i);
                if (!compareTwoCells(cell1, cell2)) {
                    equalRows = false;
                    console.log("not equal cells");
                    break;
                }
                else {
                    console.log("Equal cells");
                }
            }
            return equalRows;
        }
        function compareTwoCells(cell1, cell2) {
            if ((cell1 == null) && (cell2 == null)) {
                return true;
            }
            else if ((cell1 == null) || (cell2 == null)) {
                return false;
            }
        }
    });
});
