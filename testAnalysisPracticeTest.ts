
const testData = {
  "success": true,
  "statusCode": 200,
  "data": {
    "_id": "69edbcb1b99ed9e07a89e353",
    "test": {
      "_id": "69ea6565b093d1667a6b38a5",
      "type": "MockAndRegular",
      "name": "Short Practice Test 01",
      "testPageTitle": null,
      "totalQuestions": 24,
      "totalMarks": 96,
      "maxDuration": 60,
      "config": {
        "isScientificCalculatorEnabled": false,
        "proctoring": false,
        "enableBasicCalculator": false,
        "isScratchPadEnabled": false
      },
      "testLeaderBoardEnable": true,
      "testSliderEnable": true,
      "solutionVisible": true,
      "resultScheduleAt": "2026-04-27T02:00:47.000Z",
      "enablePercentile": true
    },
    "testStudentMapping": {
      "_id": "69edaf49ecfe761f573ac11d",
      "endedAt": "2026-04-26T07:20:17.378Z"
    },
    "rank": 30481,
    "stateRank": null,
    "percentile": 38.43499104718997,
    "attemptFilter": [
      {
        "title": "Attempt 1 (Live)",
        "testStudentMappingId": "69edaf49ecfe761f573ac11d",
        "isSelected": true
      }
    ],
    "sections": [
      {
        "totalScore": 32,
        "userScore": 14,
        "timeTaken": 721,
        "totalQuestions": 8,
        "attemptedQuestions": 6,
        "unAttemptedQuestions": 2,
        "correctQuestions": 4,
        "inCorrectQuestions": 2,
        "accuracy": 66.67,
        "completed": 75,
        "correctScore": 16,
        "inCorrectScore": -2,
        "unAttemptedScore": 8,
        "sectionId": {
          "name": "Physics",
          "_id": "69ea659488d561562119bfef"
        },
        "rank": 8889,
        "percentile": 31.703055507422167,
        "catPercentile": null
      },
      {
        "totalScore": 32,
        "userScore": 32,
        "timeTaken": 611,
        "totalQuestions": 8,
        "attemptedQuestions": 8,
        "unAttemptedQuestions": 0,
        "correctQuestions": 8,
        "inCorrectQuestions": 0,
        "accuracy": 100,
        "completed": 100,
        "correctScore": 32,
        "inCorrectScore": 0,
        "unAttemptedScore": 0,
        "sectionId": {
          "name": "Chemistry",
          "_id": "69ea659dd8154607fec6ba10"
        },
        "rank": 565,
        "percentile": 100,
        "catPercentile": null
      },
      {
        "totalScore": 32,
        "userScore": 5,
        "timeTaken": 2070,
        "totalQuestions": 8,
        "attemptedQuestions": 5,
        "unAttemptedQuestions": 3,
        "correctQuestions": 2,
        "inCorrectQuestions": 3,
        "accuracy": 40,
        "completed": 62.5,
        "correctScore": 8,
        "inCorrectScore": -3,
        "unAttemptedScore": 12,
        "sectionId": {
          "name": "Maths",
          "_id": "69ea65a6e6b3b6946d1c4f82"
        },
        "rank": 28785,
        "percentile": 22.717033443077455,
        "catPercentile": null
      }
    ],
    "yourPerformance": {
      "totalScore": 96,
      "userScore": 51,
      "timeTaken": 3402,
      "totalQuestions": 24,
      "attemptedQuestions": 19,
      "unAttemptedQuestions": 5,
      "correctQuestions": 14,
      "inCorrectQuestions": 5,
      "accuracy": 73.68,
      "completed": 79.17,
      "correctScore": 56,
      "inCorrectScore": -5,
      "unAttemptedScore": 20,
      "isLiveAttempt": true
    },
    "topperPerformance": {
      "totalScore": 96,
      "userScore": 96,
      "timeTaken": 61,
      "totalQuestions": 24,
      "attemptedQuestions": 24,
      "unAttemptedQuestions": 0,
      "correctQuestions": 24,
      "inCorrectQuestions": 0,
      "accuracy": 100,
      "completed": 100,
      "correctScore": 96,
      "inCorrectScore": 0,
      "unAttemptedScore": 0,
      "isLiveAttempt": true,
      "studentId": {
        "_id": "67d7e51e33c6ab3b8a9a7d4f",
        "firstName": "Ingela"
      }
    },
    "averagePerformance": {
      "score": 55.81,
      "correctQuestions": 15.04,
      "inCorrectQuestions": 4.37,
      "unAttemptedQuestions": 4.59,
      "accuracy": 72.02,
      "maximumScore": 0
    },
    "infoMessage": {
      "percentile": "Percentile is calculated when you attempt the test in the live window.",
      "rank": "Rank is calculated when you attempt the test in the live window.",
      "isPredictive": false,
      "timeAnalysisEnabled": true
    },
    "enableReattempt": true,
    "enableOfflineFeedback": false,
    "freeTestCouponFound": false,
    "freeTestCouponUnlocked": false
  }
}
interface TestSubjestWiseResult {
    totalScore: number,
    userScore: number,
    timeTaken: number,
    totalQuestions: number,
    attemptedQuestions: number,
    unAttemptedQuestions: number,
    correctQuestions: number,
    inCorrectQuestions: number,
    accuracy: number,
    completed: number,
    correctScore: number,
    inCorrectScore: number,
    unAttemptedScore: number,
    sectionId: {
        name: "Physics" | "Chemistry" | "Maths",
        _id: string,
    },
    rank?: number,
    percentile?: number,
    catPercentile?: null | undefined,
}

const sectinoData = testData.data.sections;
const getRowNumber = (workbook: ExcelScript.Workbook, workSheetName: "Physics" | "Chemistry" | "Maths" | "OverAllPerformanance") :number=>{
    const workSheet = workbook.getWorksheet(workSheetName)
    let rowNumber = 1
    let columnNumber = 0
    let cell = workSheet.getCell(rowNumber , columnNumber)
    while(  cell.getValue() !== ""){
        rowNumber ++
        cell  = workSheet.getCell(rowNumber, columnNumber)
    }
    console.log(rowNumber)
    return rowNumber
}
// handle worksheet fillment
const handleWorksheetNumberFill = (workbook: ExcelScript.Workbook, data:TestSubjestWiseResult[])=>{
    for(let i = 0 ; i < data.length; i++){
        const workSheet = workbook.getWorksheet(data[i].sectionId.name);
        const rowNumber = getRowNumber(workbook,data[i].sectionId.name )
        let cell = workSheet.getCell(rowNumber, 0)
        const valuesOfData: [string | null | object] = Object.values(data[i]) as [string | null | object]
        // valuesOfData.forEach((data)=> dtaa.dddd)
        // filling date and testName 
        cell.select()
        cell.setValue(new Date(testData.data.testStudentMapping.endedAt).toLocaleDateString())
        // finnling test name
        cell = workSheet.getCell(rowNumber, 1)
        cell.select()
        cell.setValue(testData.data.test.name)
        // filus the value
        let editableColumn = 2
        for(let j =0 ;j < valuesOfData.length; j++){
            cell = workSheet.getCell(rowNumber , editableColumn)
            editableColumn++;
            if(typeof valuesOfData[j] === "object" ) {
                editableColumn--;
                continue
            }
            cell.select()
            cell.setValue(valuesOfData[j])
        }
    }
   
}
const handleOverAllPermance = (workBook:ExcelScript.Workbook)=>{
    const workSheet = workBook.getWorksheet("OverAllPerformanance")
    const rowNumber = getRowNumber(workBook, "OverAllPerformanance")
    let cell = workSheet.getCell(rowNumber, 0);
    const values:[null | number | object]  = Object.values(testData.data.yourPerformance) as [null | number | object]
    cell.select()
    cell.setValue(new Date(testData.data.testStudentMapping.endedAt).toLocaleDateString())
    cell = workSheet.getCell(rowNumber, 1)
    cell.select()
    cell.setValue(testData.data.test.name)
    let editableColumn = 2;

    for(let i = 0 ; i < values.length; i++  ){
        cell = workSheet.getCell(rowNumber, editableColumn)
        editableColumn++;
        cell.select()
        if( typeof values[i] === "boolean"){
            editableColumn--;
            continue
        }
        cell.setValue(values[i])
    }
    cell = workSheet.getCell(rowNumber, editableColumn )
    cell.select()
    cell.setValue(testData.data.rank);


    cell = workSheet.getCell(rowNumber, editableColumn +1)
    cell.select()
    cell.setValue(testData.data.percentile);
}
function main(workbook: ExcelScript.Workbook) {
    handleWorksheetNumberFill(workbook, sectinoData as TestSubjestWiseResult[])
    handleOverAllPermance(workbook)
}
