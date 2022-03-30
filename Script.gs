let SHEET;
function unZipIt() 
{
  SHEET = SpreadsheetApp.openById("1iZ85hnkPKaZyLoTq8P6ZfxDiICRyBir5xBAhj9kh1SM").insertSheet();
  let theFolder = DriveApp.getFolderById('1Jl5fVc3hOsUsKXqVGN4agrYGnsAop4JF');
  let theFile = latestFile(theFolder);
  let fileBlob = theFile.getBlob();
  fileBlob.setContentType("application/zip");
  let unZippedfile = Utilities.unzip(fileBlob);
  let folderName = theFile.getName().replace(/\.[^/.]+$/, "");
  let newFolder = theFolder.createFolder(folderName);
  unZippedfile.forEach( file =>
  {
    newFolder.createFile(file);
  });
  singleOrMultiple(newFolder);
}

function singleOrMultiple(folder)
{
  let files = folder.getFiles();

  let theFile = files.next();
  if(theFile.getMimeType() === "application/zip" || theFile.getMimeType() === "application/x-zip-compressed") getZips(folder);
  else mergeCSVToSpreadsheet(folder);
}

function getZips(folder)
{
  let files = folder.getFiles();

  while(files.hasNext())
  {
    let theFile = files.next();
    let fileBlob = theFile.getBlob();
    fileBlob.setContentType("application/zip");
    let unZippedfile = Utilities.unzip(fileBlob);
    let folderName = theFile.getName().replace(/\.[^/.]+$/, "");
    let newFolder = folder.createFolder(folderName);
    unZippedfile.forEach( file =>
    {
      newFolder.createFile(file);
    });
    mergeCSVToSpreadsheet(newFolder);
  }
}

function latestFile(folder)
{
  let arrayFileDates = [],
      fileDate, 
      newestDate,
      newestFileID,
      objFilesByDate={};
  
  // Files are application/x-zip-compressed if downloadeded directly from export
  let files = folder.getFilesByType("application/zip");
  let result = [];
  while (files.hasNext()) {
    let file = files.next();
    fileDate = file.getLastUpdated();
    objFilesByDate[fileDate] = file.getId();

    arrayFileDates.push(file.getLastUpdated());
  }

  if (arrayFileDates.length === 0) {
    return;
  }

  arrayFileDates.sort(function(a,b){return b-a});

  newestDate = arrayFileDates[0];
  newestFileID = objFilesByDate[newestDate];

  return DriveApp.getFileById(newestFileID);
}

function mergeCSVToSpreadsheet(folder)
{
  let csvData = [];
  let fileNames = [];
  let files = folder.getFiles();

  while(files.hasNext())
  {
    let currentFile = files.next();
    fileNames.push(currentFile.getName());
    let csvStringData = currentFile.getBlob().getDataAsString();
    while(csvStringData.indexOf(";")>-1)
    {
      csvStringData = csvStringData.replace(";",",");
    }
    csvData.push(Utilities.parseCsv(csvStringData));
  }

  let startColumn;
  csvData.forEach((csv, index) =>
  {
    startColumn = SHEET.getLastColumn();
    Logger.log(startColumn);
    if(startColumn > 1) startColumn += 2;
    else startColumn++;
    SHEET.getRange(1, startColumn,1, 1).setValue(fileNames[index]);
    SHEET.getRange(2, startColumn, csv.length, csv[0].length).setValues(csv);
  })
}
