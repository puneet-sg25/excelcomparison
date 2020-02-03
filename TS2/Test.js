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
    for (var i = 1; i < sheetNames1.length; ++i) {
        for (var a = 2; a <= W1WorkSheets[i].rowCount; a++) {
            for (var b = 2; b <= W1WorkSheets[i].columnCount; b++) {
                if (W1WorkSheets[i].getRow(a).getCell(b).value == null) {
                    continue;
                }
                if (sheetNames1[i] != W1WorkSheets[0].getRow(a).getCell(b).value || sheetNames1[i] != W1WorkSheets[i].getRow(a).getCell(b).value) {
                    // console.log("values equal");
                    dub = 1;
                    var error = sheetNames1[0] + ' and ' + sheetNames1[i] + ' are diffrent at row ' + (a) + ', column ' + (b);
                    arrayOfEror.push(error);
                }
            }
        }
    }
    if (dub == 1) {
        for (var i_2 = 0; i_2 < arrayOfEror.length; i_2++) {
            console.log(arrayOfEror[i_2]);
        }
    }
    else {
        console.log("All sheets are ok");
    }
}
;
