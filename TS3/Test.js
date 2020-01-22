var Excel = require('exceljs');
var wb = new Excel.Workbook();
var countSheet1 = new Set();
var countSheet2 = new Set();
wb.xlsx.readFile("./Sample.xlsx").then(function (worksheet1, sheetId1) {
    var sh1 = wb.getWorksheet(sheetId1);
    //console.log("sheetId1 ==" +worksheet1);
    countSheet1.add(sheetId1);
    wb.xlsx.readFile("./SampleAgain.xlsx").then(function (worksheet2, sheetId2) {
        var sh2 = wb.getWorksheet(sheetId2);
        // console.log("sheetId2 ==" +worksheet2);
        countSheet2.add(sheetId2);
        if (countSheet1.size != countSheet2.size) {
            console.log("Different file Name");
        }
        var flag = 0;
        var dub = 0;
        var arrayOfEror = [];
        if (sh1.rowCount == sh2.rowCount && sh1.columnCount == sh2.columnCount) {
            for (var i = 1; i <= sh1.rowCount; i++) {
                for (var j = 1; j <= sh1.columnCount; j++) {
                    if (sh1.getRow(i).getCell(j).value != sh2.getRow(i).getCell(j).value) {
                        flag = 1;
                        arrayOfEror.push((i) + ',' + (j));
                    }
                }
            }
        }
        else {
            if (sh1.rowCount > sh2.rowCount) {
                console.log("missing rows in second");
            }
            else if (sh1.rowCount < sh2.rowCount) {
                console.log("missing rows in first");
            }
            if (sh1.columnCount > sh2.columnCount) {
                console.log("missing columns in second");
            }
            else if (sh1.columnCount < sh2.columnCount) {
                console.log("missing columns in first");
            }
            dub = 1;
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
    });
});
