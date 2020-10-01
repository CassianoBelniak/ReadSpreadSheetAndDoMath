// Const declarations
const xlsx = require('xlsx');
const FILE_PATH = "test-fil.xlsx"


// Program main execution
var workbook = readFile();
var sheet = getFirstSheet(workbook);


// Function Declaration
function readFile(){
    try {
        return xlsx.readFile(FILE_PATH);
    } catch (e) {
        throw `Error reading ${FILE_PATH} file` 
    }
}

function getFirstSheet(workbook){
    if (workbook.Sheets.length === 0)
        throw "The file does not contains any sheet"
    return workbook.Sheets[workbook.SheetNames[0]];
}