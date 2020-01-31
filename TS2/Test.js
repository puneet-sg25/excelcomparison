"use strict";
exports.__esModule = true;
var Excel = require('exceljs');
var XLSX = require("xlsx");
var wb = new Excel.Workbook();
var wb1 = XLSX.readFile("./NewData.xlsx");
var sheetNames1 = wb1.SheetNames;
var W1WorkSheets = [];
wb.xlsx.readFile("./NewData.xlsx").then(function () {
    var sheetName1 = sheetNames1[0];
    W1WorkSheets.push(wb.getWorksheet(sheetName1));
    for (var i = 1; i < sheetNames1.length; ++i) {
        var sheetName = sheetNames1[i];
        W1WorkSheets.push(wb.getWorksheet(sheetName));
    }
    fi();
});
function fi() {
    var flag = 0;
    var dub = 0;
    var arrayOfEror = [];
    for (var i = 1; i < sheetNames1.length; ++i) {
        for (var a = 2; a <= W1WorkSheets[i].rowCount; a++) {
            if (W1WorkSheets[0].getRow(a).getCell(1).value != W1WorkSheets[i].getRow(a).getCell(1).value) {
                flag = 1;
                var error = sheetNames1[0] + ' and ' + sheetNames1[i] + ' are diffrent at row ' + (a) + ', column ' + (1);
                arrayOfEror.push(error);
            }
        }
    }
    if (flag == 1) {
        for (var i_1 = 0; i_1 < arrayOfEror.length; i_1++) {
            console.log(arrayOfEror[i_1]);
        }
    }
    if (flag == 0) {
        console.log("All TC's are same in Business Flow and its keyword's sheets");
    }
}
;
/*
     if(sh1.rowCount == sh2.rowCount == sh3.rowCount == sh4.rowCount == sh5.rowCount == sh6.rowCount && sh1.columnCount == sh2.columnCount == sh3.columnCount == sh4.columnCount== sh5.columnCount == sh6.columnCount){
            for(let i=2;i<=sh1.rowCount;i++){
                for(let j=1;j<=sh1.columnCount;j++){
                    if(sh1.getRow(i).getCell(j).value == sh2.getRow(i).getCell(j).value){
                        console.log(sh2.getRow(i).getCell(j).value);
                        if(sh2.getRow(i).getCell(j).value == null){
             
                            j++;
                        }
                    }
                }
            }

            for(let i=2;i<=sh1.rowCount;i++){
                for(let j=1;j<=sh1.columnCount;j++){
                    if(sh1.getRow(i).getCell(j).value == sh3.getRow(i).getCell(j).value){
                        console.log(sh3.getRow(i).getCell(j).value);
                        if(sh3.getRow(i).getCell(j).value == null){
             
                            j++;
                        }
                    }
                }
            }
            
             for(let i=2;i<=sh1.rowCount;i++){
                for(let j=1;j<=sh1.columnCount;j++){
                    if(sh1.getRow(i).getCell(j).value == sh4.getRow(i).getCell(j).value){
                        console.log(sh4.getRow(i).getCell(j).value);
                        if(sh4.getRow(i).getCell(j).value == null){
             
                            j++;
                        }
                    }
                }
            }

             for(let i=2;i<=sh1.rowCount;i++){
                for(let j=1;j<=sh1.columnCount;j++){
                    if(sh1.getRow(i).getCell(j).value == sh5.getRow(i).getCell(j).value){
                        console.log(sh5.getRow(i).getCell(j).value);
                        if(sh5.getRow(i).getCell(j).value == null){
             
                            j++;
                        }
                    }
                }
            }
            
             for(let i=2;i<=sh1.rowCount;i++){
                for(let j=1;j<=sh1.columnCount;j++){
                    if(sh1.getRow(i).getCell(j).value == sh6.getRow(i).getCell(j).value){
                        console.log(sh6.getRow(i).getCell(j).value);
                        if(sh6.getRow(i).getCell(j).value == null){
             
                            j++;
                        }
                    }
                }
            }
        }*/
