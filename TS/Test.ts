import Excel = require('exceljs');
import * as XLSX from 'xlsx';
var wb = new Excel.Workbook();

var workbook = XLSX.readFile("./sample.xlsx");
   
    var sheets = workbook.Sheets;
    var sheetNames = workbook.SheetNames;

    for (var i = 0; i < workbook.SheetNames.length; ++i) {
       var sheet = workbook.Sheets[workbook.SheetNames[i]];
       console.log(workbook.SheetNames[i]);
       }
       wb.xlsx.readFile("./sample.xlsx").then(function(){
       var sh = wb.getWorksheet("Sheet1");
       var sh1 = wb.getWorksheet("Sheet2");

        for (let i = 1; i <= sh.rowCount; i++) { 
            console.log(sh.getRow(i).getCell(1).value);
            console.log(sh.getRow(i).getCell(2).value);
        }
        for (let i = 1; i <= sh1.rowCount; i++) { 
            console.log(sh1.getRow(i).getCell(1).value);
            console.log(sh1.getRow(i).getCell(2).value);
        }
});