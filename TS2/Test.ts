var Excel = require('exceljs');
import * as XLSX from 'xlsx';
var wb = new Excel.Workbook();

var wb1: XLSX.WorkBook = XLSX.readFile("./NewData.xlsx");
 var sheetNames1 = wb1.SheetNames;
var W1WorkSheets = [];
 
wb.xlsx.readFile("./NewData.xlsx").then(function(){
    let sheetName1 = sheetNames1[0]
    W1WorkSheets.push(wb.getWorksheet(sheetName1))
   
    for (var i = 1; i < sheetNames1.length; ++i) {
                let sheetName = sheetNames1[i]
                W1WorkSheets.push(wb.getWorksheet(sheetName))
            }

            fi();
});

function fi() {

    var flag = 0;
    var arrayOfEror = [];      

            for(var i=1;i<sheetNames1.length;++i){
                for(var a=2;a<=W1WorkSheets[i].rowCount;a++){
                  if(W1WorkSheets[0].getRow(a).getCell(1).value != W1WorkSheets[i].getRow(a).getCell(1).value){
                       flag = 1;
                       let error = sheetNames1[0] + ' and ' + sheetNames1[i] +' are diffrent at row ' +(a) + ', column ' + (1)
                       arrayOfEror.push(error);
                }
            }
        }
     if(flag == 1){
        for(let i=0;i<arrayOfEror.length;i++){
            console.log(arrayOfEror[i]);
        }
    }
    if(flag == 0){
        console.log("All TC's are same in Business Flow and its keyword's sheets");
    }
};