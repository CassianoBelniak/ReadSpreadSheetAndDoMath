// Constants declarations
const xlsx = require('xlsx');
const readline = require("readline");
const { resolve } = require('path');

const INSCRITION_COLUMN = "A";
const NAME_COLUMN = "B";
const ABSENCES_COLUMN = "C";
const P1_GRADE_COLUMN = "D";
const P2_GRADE_COLUMN = "E";
const P3_GRADE_COLUMN = "F";
const STARTING_INDEX = 4;
const ENDING_INDEX = 28;
const SITUATION_COLUMN = "G";
const GRADE_FOR_APPROVATION_COLUMN = "H";

const TOTAL_CLASSES = 60;

const SITUATION = {'APPROVED': 0, 'DISAPPROVED_BY_GRADE':1, 'DISAPPROVED_BY_FREQUENCY':2, 'FINAL': 3};
const SITUATION_LABEL = {0:'Aprovado', 1: 'Reprovado por nota', 2: 'Reprovado por faltas', 3: 'Final'};

// Program main execution
async function main(){
    try {
        console.log('Iniciando programa');
        var filePath = await getFilename();
        console.log(`Lendo arquivo ${filePath}`);
        var workbook = readFile(filePath);
        var sheet = getFirstSheet(workbook);
        console.log('Calculando resultados...');
        var results = getResults(sheet);
        showResults(results);
        await updateSpreadSheetWithResults(filePath, workbook, sheet, results);
        console.log('Pronto!');
    } catch (e) {
        console.error(e);
    }
}
main();


// Function Declaration
async function getFilename(){
    var fileName = await readLineAsync('Entre com o caminho do arquivo: ');
    return fileName;
}

function readFile(filePath){
    try {
        return xlsx.readFile(filePath);
    } catch (e) {
        throw `Error reading ${filePath} file`; 
    }
}

function getFirstSheet(workbook){
    if (workbook.Sheets.length === 0)
        throw "The file does not contains any sheet";
    return workbook.Sheets[workbook.SheetNames[0]];
}

function getCellValue(sheet, cell){
    if (!sheet[cell])
        throw `The cell ${cell} does not contains any value`;
    return sheet[cell].v;
}

function getResults(sheet){
    var results = [];
    for (row = STARTING_INDEX; row < ENDING_INDEX; row++){
        var name = getCellValue(sheet, NAME_COLUMN + row);
        var inscritionNumber = getCellValue(sheet, INSCRITION_COLUMN + row);
        var p1 = getCellValue(sheet, P1_GRADE_COLUMN + row);
        var p2 = getCellValue(sheet, P2_GRADE_COLUMN + row);
        var p3 = getCellValue(sheet, P3_GRADE_COLUMN + row);
        var absences = getCellValue(sheet, ABSENCES_COLUMN + row);
        var averageGrade = Math.round((p1 + p2 + p3)/3);
        var situation = getSituation(averageGrade, absences);
        var finalGrade = getFinalGrade(situation, averageGrade);
        results.push({row, name, inscritionNumber, averageGrade, situation, finalGrade, absences});
    }
    return results;
}

function getSituation(averageGrade, absences){
    if (absences/TOTAL_CLASSES > 0.25)
        return SITUATION.DISAPPROVED_BY_FREQUENCY;
    if (averageGrade < 50)
        return SITUATION.DISAPPROVED_BY_GRADE;
    if (averageGrade < 70)
        return SITUATION.FINAL;
    return SITUATION.APPROVED;
}

function getFinalGrade(situation, averageGrade){
    if (situation !== SITUATION.FINAL)
        return 0;
    return Math.round(100-averageGrade); //(5 <= (m + naf)/2)
}

function showResults(results){
    console.log("Results:");
    results.forEach(result=>{
        console.log(`${pad(result.inscritionNumber.toString(),2)} - ${pad(result.name, 15)}: Situação: ${pad(SITUATION_LABEL[result.situation], 20)} - Nota para a aprovação final - ${result.finalGrade}`);
    });
    console.log();
}

function pad(string, size){
    while(string.length < size){
        string += ' ';
    }
    return string;
}

async function updateSpreadSheetWithResults(filePath, workbook, sheet, results){
    var response = ""
    while (response !== 'S' && response !== 'N'){
        response = await readLineAsync(`Salvar resultados no arquivo ${filePath}?(S/N): `);
        response = response.toUpperCase();
    }
    if (response === 'N')
        return null;
    results.forEach(result=>{
        sheet[SITUATION_COLUMN+result.row] = {"v": SITUATION_LABEL[result.situation]};
        sheet[GRADE_FOR_APPROVATION_COLUMN+result.row] = {'v': result.finalGrade};
    });
    try{
        xlsx.writeFile(workbook, filePath);
        console.log('Resultados salvos!');
    } catch (e) {
        throw `Error when saving file ${filePath}`;
    }
}

function readLineAsync(message) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
      });
    return new Promise((resolve, reject) => {
      rl.question(message, (answer) => {
        rl.close();
        resolve(answer);
      });
    });
} 