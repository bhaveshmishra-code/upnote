---
categories:
- Inbox
---
# Create Google Docs

  

```
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSheet()
  const numRows = sheet.getLastRow()
  SpreadsheetApp.getUi()
    .createMenu('Dharani')
    .addItem(`Generate Proceedings`, 'generateProceedings')
    .addItem('Generate PDFs', 'generatePDFs')
    .addToUi();
}
function generateProceedings() {
  let templateDoc = DocumentApp.openById('1_ja635TpEE_vzXhfcfcMsRzM7ZwWUL9KPS7WKFDOX3E')
  const sheet = SpreadsheetApp.getActiveSheet()
  const numRows = sheet.getLastRow()
  const numCols = sheet.getLastColumn()
  const values = sheet.getRange(1, 1, numRows, numCols).getValues()
  //Dharani Applications folder
  const dharaniFolder = DriveApp.getFolderById('1w5PtNSIj1YwAkntpIZrLmNPkzwxbdgZs')
  const temp = new Date(Date.now())
  const dateString = `${temp.getDate()}-${temp.getMonth() + 1}-${temp.getFullYear()}`
  //first row contains headers so starting iteration from 2nd row
  for (let i = 1; i < numRows; i++) {
    const proceedingNumber = values[i][0]
    const mandalName = values[i][2]
    const village = values[i][3]
    const moduleName = values[i][4]
    const applicationNumber = values[i][5]
    const applicantName = values[i][6]
    const fatherHusbandName = values[i][7]
    const residentialAddress = values[i][8]
    const surveyNumbers = values[i][9]
    const extent = values[i][10]
    const khata = values[i][11]
    const remarks = values[i][12]
    const status = values[i][13]
    console.log(applicationNumber)
    const newFile = DriveApp.getFileById(templateDoc.getId()).makeCopy().setName(applicationNumber)
    const newDoc = DocumentApp.openById(newFile.getId())
    const body = newDoc.getBody()
    body.replaceText('{PROCEEDING NUMBER}', proceedingNumber)
    body.replaceText('{DATE}', dateString)
    body.replaceText('{MANDAL NAME}', mandalName)
    //body.replaceText('', village)
    body.replaceText('{MODULE NAME}', moduleName)
    body.replaceText('{APPLICATION NUMBER}', applicationNumber)
    body.replaceText('{APPLICANT NAME}', applicantName)
    body.replaceText('{FATHER NAME}', fatherHusbandName)
    body.replaceText('{RESIDENTIAL ADDRESS}', residentialAddress)
    body.replaceText('{SURVEY NUMBER}', surveyNumbers)
    body.replaceText('{EXTENT}', extent)
    body.replaceText('{KHATA}', khata)
    body.replaceText('{REMARKS}', remarks)
    body.replaceText('{STATUS}', status)
    newFile.moveTo(dharaniFolder)
  }
}
function generatePDFs() {
  //Dharani Applications folder
  const dharaniFolder = DriveApp.getFolderById('1w5PtNSIj1YwAkntpIZrLmNPkzwxbdgZs')
  Logger.log(dharaniFolder.getName())
  //create Out folder
  let outputFolder = null
  const folderIterator = dharaniFolder.getFoldersByName('Output')
  if (folderIterator.hasNext()) {
    outputFolder = folderIterator.next()
  }
  else {
    outputFolder = dharaniFolder.createFolder('Output')
  }
  const files = dharaniFolder.getFiles()
  while (files.hasNext()) {
    const dharaniFile = files.next()
    Logger.log(dharaniFile.getName())
    const blob = dharaniFile.getBlob().getAs('application/pdf')
    outputFolder.createFile(blob)
  }
}
```