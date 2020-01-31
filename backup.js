"use strict";
exports.__esModule = true;
var Excel = require("exceljs");
var XLSX = require("xlsx");
var wb = new Excel.Workbook();
//var countSheet1 = new Set();
//var countSheet2 = new Set();
var wb1 = XLSX.readFile("./Sample.xlsx");
var wb2 = XLSX.readFile("./SampleAgain.xlsx");
var sheetNames1 = wb1.SheetNames;
var sheetNames2 = wb2.SheetNames;
if (sheetNames1.length != sheetNames2.length) {
    console.log("Different file size");
}
var W1WorkSheets = [];
var W2WorkSheets = [];
var promise1 = new Promise(function (resolve, reject) {
    wb.xlsx.readFile("./Sample.xlsx").then(function () {
        for (var i = 0; i < sheetNames1.length; ++i) {
            var sheetName = sheetNames1[i];
            console.log(sheetName);
            W1WorkSheets.push(wb.getWorksheet(sheetName));
        }
        resolve(W1WorkSheets);
    });
    return promise1;
});
var promise2 = new Promise(function (resolve, reject) {
    wb.xlsx.readFile("./SampleAgain.xlsx").then(function () {
        for (var i = 0; i < sheetNames2.length; ++i) {
            var sheetName = sheetNames2[i];
            W2WorkSheets.push(wb.getWorksheet(sheetName));
        }
        resolve(W2WorkSheets);
    });
    return promise2;
});
Promise.all([promise1, promise2]).then(function (values) {
    var flag = 0;
    var dub = 0;
    var arrayOfEror = [];
    if (sheetNames1.length == sheetNames2.length) {
        console.log("in 1 if");
        for (var i = 0; i < sheetNames1.length; ++i) {
            console.log("in 1 for");
            if (W1WorkSheets[i].rowCount == W2WorkSheets[i].rowCount && W1WorkSheets[i].columnCount == W2WorkSheets[i].columnCount) {
                console.log("in 2 if");
                for (var a = 1; a <= W1WorkSheets[i].rowCount; a++) {
                    console.log("in 2 for");
                    for (var b = 1; b <= W1WorkSheets[i].columnCount; b++) {
                        console.log("in 3 for");
                        console.log(W1WorkSheets[i].getRow(a).getCell(b).value);
                        console.log(W2WorkSheets[i].getRow(a).getCell(b).value);
                        if (W1WorkSheets[i].getRow(a).getCell(b).value != W2WorkSheets[i].getRow(a).getCell(b).value) {
                            flag = 1;
                            var error = sheetNames1[i] + ' and ' + sheetNames2[i] + ' are diffrent at row ' + (a) + ', column ' + (b);
                            arrayOfEror.push(error);
                        }
                    }
                }
            }
            else {
                if (W1WorkSheets[i].rowCount > W2WorkSheets[i].rowCount) {
                    console.log("missing rows in second");
                }
                else if (W1WorkSheets[i].rowCount < W2WorkSheets[i].rowCount) {
                    console.log("missing rows in first");
                }
                if (W1WorkSheets[i].columnCount > W2WorkSheets[i].columnCount) {
                    console.log("missing columns in second");
                }
                else if (W1WorkSheets[i].columnCount < W2WorkSheets[i].columnCount) {
                    console.log("missing columns in first");
                }
                dub = 1;
            }
        }
        if (flag == 1) {
            console.log("Different files at position -->");
            for (var i = 0; i < arrayOfEror.length; i++) {
                console.log(arrayOfEror[i]);
            }
        }
        if (dub == 0 && flag == 0) {
            console.log("same files");
        }
    }
    else {
        console.log("Different file size");
    }
});
