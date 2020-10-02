// Constants declarations
const { GoogleSpreadsheet } = require('google-spreadsheet');
const FILE_ID = "1nfo8YE9vrpARbdpk9tG7ezeDmfd1G_tK"

const CREDENTIALS_EMAIL = 'defaultconnection@spreadsheetsmaths.iam.gserviceaccount.com';
const CREDENTIALS_ID = "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQC4333yd3SBoiP1\nX78xwNbvyzXguEdwrABCatkjmY7xjdZDgVzX93Nz7slrwF3iM9JMajaUrgwYL3Dp\njryMNfw8z8AbUi7H5q9cXgAbVlwO1SQUf/cxvK2hMf1+uVeyayWjiK/FMPLs5wQi\n6Ju6eLTt7UffrINYktryiUDTnh0HnGTVAoyBLJYwcunqQujDUQrNy6nHbnv78C/I\nZT74/34mESML/PiMUbY91a2r34SScoiIcA8fWdzcwJwlQ1f5X8vfNKH5evmcoPyP\nS1xeg1Lez/J/ZyAoeERZvxjDtCqWX013Dbr01UXUN69KLWSq+PCi8NvaLuhW7MXn\n6gylwxEVAgMBAAECggEABTvEbbr+eBX+NRnJCJBINWEYDRzCXvjrh/XM4FJeRs3I\nWUBd/7ogUVGa75syPS9Q3ntqQKK9smiTZnU1NrXnhlQuQMe+jcekwrVhhOSYtg3I\nF/F4brbD5oqK/c2i4yjf16WMrkUgt16h0hgqImj83DhpnrYcQMNlgdSrWmJBOaAi\nM6BXZfEVYUd5l57TkD+11d6oPAq9Q03mSJYKB+LqBf/+K/DvVWmN6FL93CMUAMVp\njDDFsQmfu0ji9ePZCJIvYVtTyD8TZ7n7wTw3pzkd0o9zjupB80BxSelpq0rspn1f\nTV1plu1yW0cqxAZWW+6VDQeucTIIc45WRhXfO9eR+QKBgQDefFOjkZJqLULO/mnM\n4OOhHkIhOO88l4kn1ONN6jggWFbcpqhlyf0Bw1GrhpBP6Bq6LOtL/jj+qiUOJscR\nmJgSMVLrCP3tt5TIztaJuBSJoSiNgI2kKD++a2UkrFKC/YfJiKEc/VU/Ag2Nf8Rz\nAxWqAgy8jh0+ya9p52tU1QijrQKBgQDUuLdS3uiJ3MdSac6HDip6RH64uJYuLB2t\n0DUZ8lObojDILndGBmiH7PSC3qJbogp9T4ENh9Dm50j4YNPWjQs1Ixt6Nw3GEPFd\nxekgafUTULotyOM7JVHNvSGb9VrVCoFfDl+DvKTLoMvqXMy5dUE2d/Iqj2AfPf7X\nIaBO6nSQCQKBgC3pUgkq/R/T/zlf3s1cixywdc0NRrEmRDNoBxAJCVQDZslZyt5W\ndFNszumqdxVGPF2270dbSr+itMrazbGf36HBc+70iBIKFDXsGPGKfxJ3ozqwEIqT\nk7PjzZdnyA8n6mF4RGcLEBBUiB9vAkcJl+rhSWePnBFc5Unha5Cx9XpxAoGAL83z\nJOSDTbgX8yVkDGXale+eqtSQq3+ui8kmpdYXg/pHDDWlCE+YXjOaH274/a7EvLSJ\nRAkpoTqI44ifErBPvHlPS3/j0IcuNuyrH2Wwdc7GiFOE/V29rIa8btgMuaPKvxnz\nzR8vybMxIFIKkAMRzLPX8EiYSW0dQCuGYzW9TEECgYAZ3fbSbKvbgBHAc1WokZ0S\nLgoiLdV6owLvJXqkFSUsMkaIPHy846hVjz5mvmiK7HGxQayvAt8/rTnkAixpM3t4\nrmE9+JCclIw2FDLNb7pGuoFw/f7JK5B0ccbd4tJ96a6I8Dtn/+lbIRM1cjVvOKVR\nF/Nmw0RD0uHPhdiu7kxKaQ==\n-----END PRIVATE KEY-----\n".replace(/\\n/g, '\n');

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
        console.log('Starting program');
        console.log(`Reading ${FILE_ID} file`);
        var workbook = await readFile();
        var sheet = getFirstSheet(workbook);
        console.log('Calculating results...');
        var results = getResults(sheet);
        showResults(results);
        //console.log("Salving results");
        //updateSpreadSheetWithResults(workbook, sheet, results);
        console.log('Done');
    } catch (e) {
        console.error(e);
    }
}
main();

// Function Declaration
async function readFile(){
    try {
        var doc = new GoogleSpreadsheet(FILE_ID);
        await doc.useServiceAccountAuth({
            client_email: CREDENTIALS_EMAIL,
            private_key: CREDENTIALS_ID,
          });
        await doc.loadInfo();
        return doc;
    } catch (e) {
        throw `Error reading ${FILE_ID} file : ${e}`; 
    }
}

function getFirstSheet(workbook){
    if (workbook.sheetsByIndex.length === 0)
        throw "The file does not contains any sheet";
    return workbook.sheetsByIndex[0];
}

function getCellValue(sheet, cell){
    if (!sheet.getCellByA1(cell))
        throw `The cell ${cell} does not contains any value`;
    return sheet.getCellByA1(cell);
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
}

function pad(string, size){
    while(string.length < size){
        string += ' ';
    }
    return string;
}

function updateSpreadSheetWithResults(workbook, sheet, results){
    results.forEach(result=>{
        sheet[SITUATION_COLUMN+result.row] = {"v": SITUATION_LABEL[result.situation]};
        sheet[GRADE_FOR_APPROVATION_COLUMN+result.row] = {'v': result.finalGrade};
    });
    try{
        xlsx.writeFile(workbook, FILE_PATH);
    } catch (e) {
        throw `Error when saving file ${FILE_PATH}`;
    }
}