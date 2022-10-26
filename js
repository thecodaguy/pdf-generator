/*
Step 1: Read Google Sheet for results
Step 2: Create a google doc with results
Step 3: Create a PDF from that google doc
Step 4: Share that PDF with the individual as a viewer
Step 5: Log the URL of that PDF in a google sheet row
Step 6: Send an email to that person with the link to their results
*/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Create PDFs')
      .addItem('Create PDFs in Google Drive', 'createPDFsClean')
      .addToUi();
};
function makeObject(array){
  let rowsAsObjects = [];
  for (let [i,row] of Array.from(array.entries())){
    if(i>0){
      let objectRow = {};
      for (let [x,value] of Array.from(row.entries())){
        let keyValue = makeKey(array[0][x])
        objectRow[keyValue] = row[x]
        objectRow.index = i-1
        objectRow.sheetRow = i+1
      }
      rowsAsObjects.push(objectRow)
    }
  }
  return rowsAsObjects
}

function makeKey(value){
  if(/\s/.test(value)){
    let key = value.split(" ").slice(0,3).join("_").toLowerCase()
    return key
  } else return value.toLowerCase();
};
function portOverInfo (){
  let originalSheet = SpreadsheetApp.openById("16sVfp1UTY0mNTbZjgLG78NiZpGNr6J9fgFSdwCoXTZI").getSheetByName("registrants");
  let originalValues = originalSheet.getDataRange().getValues();
  console.log(originalValues)
}
function createPDFsClean() {
  //GLOBAL VARIABLES, SPREADSHEETS, AND MORE
    let hpvDocs = DriveApp.getFolderById("1Ia2tyyaOdc5Yyv5vRQWRA4N1QGrNqTzJ"); //Temporary folder titled 'HPV Docs'
    let hpvPdfs = DriveApp.getFolderById("1BflUGKYIpQVPc5RGhYzcm6WlPgF90hLM"); //Final Folder called 'HPV Pdfs'
    let docTemplate = DriveApp.getFileById("1GVGoAYtQtoe94RY3kRBjpRdv9sti7BH8caC4OGKTKzg");
    let nameSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sample Tracking").getDataRange();
    let resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results").getDataRange();
    let resultRows = resultSheet.getValues() //Just an array of results
    let nameRows = nameSheet.getValues()

  //CREATE OBJECTS FOR BOTH SPREADSHEETS
    let namesObject = makeObject(nameRows)
    //let resultObject = makeObject(resultRows)
    let targetObjects = namesObject.filter(x => x.first_name.length>1)

  //FOR EACH RESULT, FIND THE TARGET NAME ROW AND CREATE A PDF
    for (let result of targetObjects){
      //DEFINE VARIABLES AND FIND TARGET OBJECT
        let pdfName = `HPV Results: ${result.first_name} ${result.last_name}`
        if (result.first_name){ //If call ensures it only runs for actual rows with names/results in it
      //CREATE A PDF
          let personalPdf = createPDF(pdfName,result)
          console.log(`Created a pdf for ${result.first_name}`)

        //INSERT PDF LINK INTO SPREADSHEET
          result.pdf = `http://drive.google.com/uc?export=view&id=${personalPdf}`
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sample Tracking").getRange(result.sheetRow,70).setValue(result.pdf)}
    }
    
  //ERASE Google Doc files
    let files = hpvDocs.getFiles();
    while(files.hasNext()){
      files.next().setTrashed(true)
  }
}
 
 
function createPDF(pdfName,data){
        const tempDocFolder = DriveApp.getFolderById("1Ia2tyyaOdc5Yyv5vRQWRA4N1QGrNqTzJ");
        const pdfFolder = DriveApp.getFolderById("1BflUGKYIpQVPc5RGhYzcm6WlPgF90hLM")
        const tempFile = DriveApp.getFileById("1GVGoAYtQtoe94RY3kRBjpRdv9sti7BH8caC4OGKTKzg").makeCopy(tempDocFolder);
        const tempDocFile = DocumentApp.openById(tempFile.getId());
        const body = tempDocFile.getBody();
        
        //Find and replace all the information from the data spreadsheet
        let dataKeys = Object.keys(data)
        for (let dataKey of dataKeys){
          let fillWith = data[dataKey];
          if (typeof fillWith == "object"){
            let formatted = fillWith.toString().split(" ").slice(1,4).join(" ");
            fillWith = formatted
          }
          body.replaceText(`{${dataKey}}`,fillWith)
        }

  /*
        //Find and replace all the info from the result keys
        let resultKeys = Object.keys(results)
        console.log("RESULTS ARE:::")
        console.log(results)
        console.log(results.hpv_16_result)
        for (let resultKey of resultKeys){
          console.log(`Round ${counter}:: Replacing --> {${resultKey}} with ${results.resultKey}`)
          body.replaceText(`{${resultKey}}`,results[resultKey]) //This is accessing the Data freaking thing
          counter = counter+1
        }*/
        
        tempDocFile.saveAndClose();
        const pdfContentBlob = tempFile.getAs(MimeType.PDF);
        let finalPDF = pdfFolder.createFile(pdfContentBlob);
        finalPDF.setName(pdfName);
        return finalPDF.getId();
        
}
