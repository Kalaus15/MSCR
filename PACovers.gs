function newCoverSheet(pa){
  SpreadsheetApp.getUi().alert("Cover sheet generation started. You can resume working and it will open in a new window when ready.");
  if(!pa){
    var email = getUser().getEmail();
    if(testing){
      email = "wlmoran@madison.k12.wi.us";
    }
    var pa = email.substring(0,email.indexOf("@"));
  }else{
    pa = pa.trim(); 
  }
  var doc = DocumentApp.create(pa + "'s payroll - created " + Utilities.formatDate(new Date(), tz, "MM/dd/YYYY"));
  var asFile = DriveApp.getFileById(doc.getId());
  var folder = getCoverSheetFolder();
  folder.addFile(asFile); //add to folder
  DriveApp.removeFile(asFile); //remove from root
  makeCoverSheet(doc,pa,getSDsforPA(pa));
  var url = doc.getUrl();
  openLink(url,"your cover sheet");
}
function newCoverSheetOther(){
  var html = HtmlService.createHtmlOutputFromFile("CSSelect")
      .setWidth(400)
      .setHeight(155);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, "Generate Cover Sheet");
}
function getPAsForCoverSheet(){
  return getList(getListSheet().getDataRange().getDisplayValues(),listSheetPACol,true);
}
function getCoverSheetFolder(){
//  switch(env){
//    case "DEV":
//    case "TST":
//      return "";
//      break;
//    case "PRD":
      return DriveApp.getFolderById("1ZrWTuWcqMHpCKWFgOncaHFJF6CXBvn3-");
//  }
}
function getCoverSheetArchive(){
  return DriveApp.getFolderById("1hGsJnsbxA2V1QPXOfMSqxX3moAY4s589"); 
}
/**
 * 
 * 3/14/21 MDH #394 deprecate addFile/removeFile
 */
function archiveCoverSheets(){
  var folder = getCoverSheetFolder();
  var archive = getCoverSheetArchive();
  var files = folder.getFiles();
  var now = (new Date()).valueOf();
  while(files.hasNext()){
    var file = files.next();
    if(now - file.getLastUpdated().valueOf() > 2592000000){
      file.moveTo(archive); // * 3/14/21 MDH #394 deprecate addFile/removeFile
    }
  }
}
function makeCoverSheet(doc,pa,sds){
  var body = doc.getBody();
  var margin = 36; //0.5"
  body.setMarginBottom(margin);
  body.setMarginLeft(margin);
  body.setMarginRight(margin);
  body.setMarginTop(margin);  
  body.insertParagraph(0,pa + "'s Payroll for ______")
  .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  .editAsText()
  .setFontSize(24);
  var table = body.appendTable();
  var headerRow = table.appendTableRow();
  var headerColNames = ["Received Date/Time","Approver","Batch #","Released Date/Time","Notes"];
  var widths = [108,90,54,108,180];
  for(var i = 0; i < headerColNames.length; i++){
    headerRow.appendTableCell(headerColNames[i])
    .setWidth(widths[i])
    .editAsText()
    .setBold(true)
    .setBackgroundColor("#cccccc")
    .setFontSize(11);
  }
  for(var i = 0; i < sds.length; i++){
    var sd = sds[i];
    var mailbox = sd.substring(sd.indexOf("<")+1,sd.indexOf("@"));
    var row = table.appendTableRow();
    for(var j = 0; j < headerColNames.length; j++){
      var cell;
      if(j == 1){
        cell = row.appendTableCell(mailbox);
      }else{
        cell = row.appendTableCell(); 
      }
      cell.getChild(0).asParagraph().editAsText().setFontSize(11).setBold(false);
    }
  }
  doc.saveAndClose();
}
function getSDsforPA(pa){
  var hired = getHiredSheet().getDataRange().getDisplayValues();
  var active = getActiveSheet().getDataRange().getDisplayValues();
  var allStaff = hired.concat(active);
  var sds = [];
  for(var i = 1; i < allStaff.length; i++){
    var row = allStaff[i];
    if(row[PACol-1].indexOf(pa) == -1){
      continue; 
    }
    var sd = row[payrollApvsCol-1];
    if(!sd){
      sd = row[supCol-1]; 
    }
    if(sd && sds.indexOf(sd) == -1){
      sds.push(sd); 
    }
  }
  sds.sort();
  sds.push("");
  return sds;
}