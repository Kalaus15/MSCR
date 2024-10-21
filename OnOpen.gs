//Copyright (c) Kyle Egerdal 2018. All Rights reserved.

var payRollOption = "Select Payroll Approvers";
var payrollThing = "Use SDs > " + payRollOption;

function onOpen(event){
  makeMenu();
  //setFocus(); *10/15/2019 KJE #177 focusing causes problems
  displayLock();
}
function displayLock(){
  var protection = getHiredSheet().getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if(!protection){
    return; 
  }
  var dotw = (new Date()).getDay();
  if(dotw == 0 || dotw == 6){
    var words = "The SEDbots are working. The SED will be read-only until Sunday at 5 AM.";
  }else{
    var words = "The SEDbots are working. The SED will be read-only for a little while.";
  }
  SpreadsheetApp.getActive().toast(words,"🔒 SED is Read-Only",10);
}
function makeMenu(){
  var ui = SpreadsheetApp.getUi();
  if(iAmAPA() || iAmAMaster()){
    ui.createMenu("PAs")
    .addItem("Create Agreement(s)", "makeAgreements") //Agreements.gs
    .addItem("Recalculate \"Age At Start\"", "reAgeAtStart")
    .addItem("Make Cover Sheet (for me)", "newCoverSheet")
    .addItem("Make Cover Sheet (for someone else)", "newCoverSheetOther")
    .addItem("Open Cover Sheet folder", "openCoverFolder")
    .addSeparator()
    .addItem("Check for Files (all my staff)", "newCheckForFilesManual")//FileChecks.gs
    .addItem("Check for Files (selected staff)", "newCheckForFilesSelected")//FileChecks.gs
    .addSeparator()
    .addItem("Check for Completion (all my staff)", "toPAManual")//PAEmails.gs
    .addItem("Check for Completion (selected staff)", "toPASelected")//PAEmails.gs
    .addSeparator()
    .addItem("Draft Cleared Email (all my staff)", "draftCleared")
    .addItem("Draft Cleared Email (selected staff)", "draftClearedSelected")
    .addSeparator()
    .addItem("Send all e-mails in my drafts folder","sendAllDrafts")
    .addSeparator()
    .addItem("Turn on SEDbot", "setUpMyTriggers")//PAEmailMgr.gs
    .addItem("Turn off SEDbot", "deleteMyTriggers")
    .addSeparator()
    .addItem("Magic Form","openMagic")
    .addItem("Handbook","openHandbook")
    .addItem("Hoyt Paperwork Appointments","openHoytApts")
    .addToUi();
  }
  ui.createMenu("SDs")
  .addItem("Get staff list","getStaffList") //*7/20/20 KJE #152
  .addItem("Can this staff log time?", "canLogTime")
  //.addItem("Grandfather this staff's pay rate", "grandfatherPayRate") //*10/8/24 KJE #549 remove this option
  .addSeparator()
  .addItem("Duplicate Employee(s)", "duplicateLine")
  .addSeparator()
  .addItem("Get Timesheets Here!", "openGetTimesheets")
  .addItem("Fix-It Form", "openFixIt")
  .addItem("FAQ", "openFAQ")
  .addItem("Statistics", "openStats")
  //  .addItem(payRollOption, "payrollSidebar")
  .addToUi();
  if(iAmAMaster()){
    ui.createMenu("Masters")
    .addItem("Lock Hired Sheet", "lockSED") //*7/20/20 KJE #266
    .addItem("Unlock Hired Sheet", "unlockSED") //*7/20/20 KJE #266
    .addSeparator() //*7/20/20 KJE #266
    .addItem("Push MUNIS Updates", "pushMUNISUpdates")
    //.addToUi();
  //}
  //if(Session.getActiveUser().toString() == getSS().getOwner().toString()){
    //ui.createMenu("Owner")
    .addSeparator()
    .addItem("Update PAs", "openUpdateForm")
    .addItem("Update SDs", "openUpdateForm")
    .addItem("Update Masters", "openUpdateForm")
    .addItem("Update Admins", "openUpdateForm")
    .addItem("Update End Dates", "openUpdateForm") //*10/8/2019 KJE *176
    .addSeparator()
    .addItem("Format Hired Sheet", "formatHired")
    .addItem("Format Hired Sheet (starting at...)", "formatHiredFrom")
    .addItem("Format Staff Lists Sheet", "formatStaffLists")
    .addSeparator()
    .addItem("Count Early Receipts", "earlyReceiptCount")
    .addItem("E-mail Me Site Totals", "sendSiteTotals")
    .addSeparator()
    .addItem("Schedule Mass Upload", "scheduleMassUpload")
    .addItem("Cancel Mass Upload", "cancelMassUpload")
    .addItem("Open Mass Upload Folder", "openMUFolder")
    .addItem("Mass Upload Instructions", "massUploadInstr")
    .addSeparator()
    .addItem("Open Canned Responses Folder", "openCannedResponsesFolder")
    .addSeparator()
    .addItem("Open ToDo List", "openToDoList")
    .addSeparator()
    .addItem("Turn On SEDbot Reminders","turnResetReminderOn")
    .addItem("Turn Off SEDbot Reminders","turnResetReminderOff")
    .addToUi();
  }
//  var email = getUser().getEmail().toUpperCase();
//  if(email.indexOf("KJEGERDAL") != -1 || email.indexOf("EMPAPE") != -1){
//    ui.createMenu("Erica")
//    .addItem("Draft Reminder E-mails", "ericaReminders")
//    .addItem("Delete all drafts", "tempDeleteAllDrafts")
//    .addItem("Send all e-mails in my drafts folder","sendAllDrafts")
//    .addToUi();
//  }
}
function getStaffList(){ //*7/20/20 KJE #152
  var sheet = SpreadsheetApp.getActiveSheet();
  if(sheet.getName() != SARLib.listSheetName){
    SpreadsheetApp.getUi().alert("To generate a list of staff, first go the to the \"Staff Lists\" tab and select one or multiple sites." + 
                                "\n\nTo select multiple sites, FIRST, select one site. Then, hold \"ctrl\" and select others." +
                                "\n\nAfter you've selected the sites you want, choose this menu option again."); 
    return;
  }
  //get site list
  var selection = sheet.getSelection();
  var rangeList = selection.getActiveRangeList();
  var ranges = rangeList.getRanges();
  var sites = [];
  for(var i = 0; i < ranges.length; i++){
    var range = ranges[i];
    var values = range.getValues();
    for(var j = 0; j < values.length; j++){
      var value = values[j][0];
      if(sites.indexOf(value) == -1){
        sites.push(value); 
      }
    }
  }
  //get staff at that site
  getHiredSheet();
  var hiredVals = hiredSheet.getRange(1, 1, hiredSheet.getMaxRows(), SARLib.lastKeepCol).getDisplayValues();
  var activeVals = getActiveSheet().getDataRange().getDisplayValues();
  var header = [activeVals[0]];
  Logger.log(header);
  var both = hiredVals.concat(activeVals);
  var bNumList = [];
  for(var i = 0; i < both.length; i++){
    var row = both[i];
    var bNum = row[SARLib.bNumCol-1];
    if(sites.indexOf(row[SARLib.siteCol-1]) == -1 || (bNum !== "NEW" && bNumList.indexOf(bNum) != -1)){ //*4/24/23 KJE #460 allow multiple "NEW" bnums
      both.splice(i, 1);
      i--;
    }else{
      bNumList.push(bNum);
    }
  }
  if(both.length == 0){
    SpreadsheetApp.getUi().alert("Couldn't find any staff at the sites below:\n\n" + sites.join(", "));
    return;
  }
  //make sheet
  var newSS = SpreadsheetApp.create("Staff List: " + (sites.length > 3 ? sites[0] + ", " + sites[1] + ", " + sites[2]  + ", ..." : sites.join(", ")) + " (" + Utilities.formatDate(new Date(), tz, "MM/dd/YY") + ")");
  var newSht = newSS.getSheets()[0];
  both = header.concat(both);
  Logger.log(both);
  var range = newSht.getRange(1, 1, both.length, both[0].length);
  range.setValues(both);
  newSht.getRange(1, SARLib.dateAddedCol).setValue("Date Added or Marked Active");
  openLink(newSS.getUrl(), "your staff list...");
}
function lockSED(){ //*7/20/20 KJE #266
  if(SARLib.SARWriteLock("mastersMenu",env,getHiredSheet(),getMasters(getListSheet().getDataRange().getValues()))){
    SpreadsheetApp.getUi().alert("Hired sheet locked!\n\nOnly the owner and those under the \"List of Masters\" column in the \"Staff Lists\" tab can edit it now.");
  }else{
    SpreadsheetApp.getUi().alert("The Hired sheet is already locked.");
  }
}
function unlockSED(){ //*7/20/20 KJE #266
  SARLib.SARWriteUnlock(env,getHiredSheet());
  SpreadsheetApp.getUi().alert("Hired sheet UN-locked!\n\nAnyone with edit access to the Spreadsheet can edit the Hired sheet now.");
}
function openUpdateForm(){
  openLink("https://docs.google.com/forms/d/e/1FAIpQLSfoK9ag3km9AR-NzL9rE-BfSbwoN0gt_AhEaUUE_aMbOX9RxA/viewform","SED update form");
}
function ownerPAUpdate(){
  getSSByID(); //*3/4/20 KJE #216 run from form
  updatePAs(getListSheet().getDataRange().getValues());
}
function ownerSDUpdate(){
  getSSByID(); //*3/4/20 KJE #216 run from form
  updateSDs(getListSheet().getDataRange().getValues());
}
function ownerMasterUpdate(){
  getSSByID(); //*3/4/20 KJE #216 run from form
  updateMasters(getListSheet().getDataRange().getValues());
}
function ownerAdminUpdate(){
  getSSByID(); //*3/4/20 KJE #216 run from form
  updateAdmins(getListSheet().getDataRange().getValues());
}
//*10/8/2019 KJE *176
function ownerEndDatesUpdate(){
  getSSByID();
  //makePleaseWait(); *3/4/20 KJE #216 run from form
  //format and sort
  var sheet = getListSheet();
  var dateRange = sheet.getRange(SARLib.listSheetLastHeader+1,
                                 SARLib.listSheetEndDatesCol,
                                 sheet.getMaxRows()-1-SARLib.listSheetLastHeader,
    1);
  dateRange.setNumberFormat("M/d/YYYY");
  dateRange.sort(SARLib.listSheetEndDatesCol);
  //apply validation
  var hired = getHiredSheet();
  var applyTo = hired.getRange(SARLib.hiredSheetLastHeader+1,
                               SARLib.endDateCol,
                               hired.getMaxRows()-1-SARLib.hiredSheetLastHeader,
    1);
  var validCol = getANotation(SARLib.listSheetEndDatesCol);
  var validRange = listSheet.getRange(validCol+(listSheetLastHeader+1)+":"+validCol+sheet.getMaxRows());
  var endDateValidation = SpreadsheetApp.newDataValidation().requireValueInRange(validRange, true).setAllowInvalid(false).build();
  applyTo.setDataValidation(endDateValidation);
  //grandfather existing dates
  var dates = dateRange.getDisplayValues().reduce(function(ac,cur,dx,ar){
    if(cur[0]){
      ac.push(cur[0]);
    }
    return ac;
  },[]);
  var vals = applyTo.getDisplayValues();
  for(var i = 0; i < vals.length; i++){
    var val = vals[i][0];
    if(val && dates.indexOf(val) == -1){
      hired.getRange(i+hiredSheetLastHeader+1,SARLib.endDateCol).clearDataValidations();
    }
  }
  //closePleaseWait(); *3/4/20 KJE #216 run from form
}
function setFocus(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if(sheet.getName() == hiredSheetName){ //might have changed sheets while loading
    ss.setActiveRange(ss.getRange(colLetter(SARLib.lastNameCol-1) + sheet.getLastRow()));
  }
}
var magicFormUrl = "https://docs.google.com/forms/d/e/1FAIpQLSf27tpt_BtP_W-lMnvGyAfcNBeEWVQZx-iwZhoBF9O-UsZUEw/viewform"
function openMagic(){
  openLink(magicFormUrl,"Magic Form");
}
function openHandbook(){
  openLink("https://docs.google.com/document/d/19VMs-5gddVLHXeTNk2H5MK5C0ARC7TvpL3G-P3ge8-Q/","Handbook");
}
function openHoytApts(){
  openLink("https://docs.google.com/spreadsheets/d/1QuZ2wKXghQiD1OlQIZHdgVnCt525lkFxqPLqIWLezok/","Hoyt Paperwork Appointments");
}
function openGetTimesheets(){
  openLink("https://docs.google.com/forms/d/e/1FAIpQLSeIyRY3uTA6SToXEzXOvT6aPxuqQZSusVBYT4APZG4dNHOezA/viewform","Get Timesheets Here!");
}
function openFixIt(){
  openLink("https://docs.google.com/forms/d/e/1FAIpQLSdWg1kdqBfp75yE25eh2DYY0XgoguYi7fmLNHiQY-uHRv6Zig/viewform","Fix-It Form");
}
function openFAQ(){
  openLink(SARLib.faq,"FAQ");
}
function openStats(){
  openLink("https://docs.google.com/spreadsheets/d/1Lyz13vwTIbnU7RoOmXT0DqypQhj-o-AuH1Qzxg3F-Ho/","System Statistics");
}
function openMUFolder(){
  openLink("https://drive.google.com/drive/u/0/folders/1A-CTUY0M8GhaXtxrX-HwvZO89Pi0ziBL","Mass Upload folder");
}
function openCannedResponsesFolder(){
  openLink(getCannedResponsesFolder().getUrl(),"Canned Responses folder"); 
}
function openToDoList(){
  openLink("https://docs.google.com/spreadsheets/d/1viAQ-3o5jMQaWYXo8suHqSOxzEhDZLHZuMgND0juYF0","ToDo List"); 
}
function openCoverFolder(){
  openLink(getCoverSheetFolder().getUrl(),"Cover Sheet folder");
}
function openLink(link,name){
  var htmlOutput = HtmlService
    .createHtmlOutput("<script>window.open(\"" + link + "\")</script>")
    .setHeight(30)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Opening " + name + "...");
  Utilities.sleep(2000); //wait long enough for window to open
  htmlOutput = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>").setHeight(30);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Opening " + name + "...");
}
function sendAllDrafts(){
  getUi();
  if((new Date()).getDay() == 1){
    ui.alert("This option can't be used on Mondays. SEDbot needs to save its strength for sending receipts out.\n\nSorry about that.");
    return;
  }
  var resp = ui.alert("Send all Drafts",
                      "This will send every single e-mail in your drafts folder,\nEVEN IF IT WASN'T GENERATED BY SEDbot.\n\n" +
                      "This includes any e-mails you are currently working on,\nbut haven't sent yet.\n\n" +
                      "It CANNOT be undone.\n\n"+
                      "Are you ABSOLUTELY SURE you want to do this?",ui.ButtonSet.YES_NO);
  if(resp != ui.Button.YES){
    return;
  }
  var failure = false;
  makePleaseWait();
  var drafts = GmailApp.getDrafts();
//  if(drafts.length > 50){
//    ui.alert("Sorry, but you can't use this to send more than 50 e-mails at a time.");
//    return;
//  }
  for(var i = 0; i < drafts.length; i++){
    try{
      drafts[i].send();
    }catch(e){
      if(e.message.indexOf("Recipient address required") != -1 || e.message.indexOf("Gmail operation not allowed.") != -1){
        failure = true;
      }else{
        throw e; 
      }
    }
  }
  closePleaseWait();
  if(failure && GmailApp.getDrafts().length > 0){
    ui.alert("At least one e-mail failed to send or errored when sent. This could be because it didn't have a \"To\" address, or maybe Gmail was being fiesty." +
             "\n\nYou can either try this menu option again, or send it/them by hand. Or, maybe it was sent, and Gmail errored for no reason at all. Check your drafts folder to be sure."); 
  }
}
function iAmAPA(){
  email = getUser().getEmail().toUpperCase();
  vals = getListSheet().getDataRange().getValues();
  if(getPAs(vals).toString().toUpperCase().indexOf(email) != -1){
    return true; 
  }
  return false;
}
//assumes: email, vals (from iAmAPA)
function iAmAMaster(){
  if(getMasters(vals).toString().toUpperCase().indexOf(email) != -1){
    return true; 
  }
  return false;
}
function pushMUNISUpdates(){
  makePleaseWait();
  getMUNISSheet();
  sortMUNIS();
  updateSiteLists();
  updateListSheet();
  closePleaseWait();
}

//assumes: munisSheet
function sortMUNIS(){
  var range = munisSheet.getRange(SARLib.munisSheetLastHeader+1,1, munisSheet.getLastRow()-SARLib.munisSheetLastHeader-1, munisSheet.getMaxColumns()); //*1/24/21 #296 use getMaxColumns
  range.sort([1,2,4]);
}
function updateSiteLists(){
  getHiredSheet();
  var range = hiredSheet.getRange(hiredSheetLastHeader+1,siteCol,hiredSheet.getLastRow()-hiredSheetLastHeader+1,1);
  var validCol = getANotation(munisSheetSiteCol);
  var munisVals = munisSheet.getDataRange().getDisplayValues();
  var lastAll = getMUNISLastAll(munisVals);
  var validRange = munisSheet.getRange(validCol+(lastAll+1)+":"+validCol+munisSheet.getLastRow());
  range.setDataValidation(SpreadsheetApp.newDataValidation()
                                             .requireValueInRange(validRange, true)
                                             .setAllowInvalid(false).build()).setHorizontalAlignment("left");
}

//assumes: munisSheet
function updateListSheet(){
  getListSheet();
  var listRange = listSheet.getDataRange();
  var listVals = listRange.getValues();
  var listValsOld = listRange.getValues();
  var munisRange = munisSheet.getDataRange();
  var munisVals = munisRange.getValues();
  var munisColors = munisRange.getBackgrounds();
  var listValSiteRow = listSheetLastHeader;
  var listValProgRow = listSheetLastHeader;
  var lastSite,lastProg;
  //get new values
  for(var i = SARLib.munisSheetLastHeader; i < munisVals.length; i++){ //*6/10/2020 KJE #259 was munisSheetLastAll even though deprecated
    if(munisVals[i][munisSheetSiteCol-1] === "!All"){
      continue; 
    }
    if(munisIgnoreColors.indexOf(munisColors[i][munisSheetSiteCol-1]) != -1){continue};
    var thisSite = munisVals[i][munisSheetSiteCol-1];
    if(thisSite != lastSite){
      if(listVals.length <= listValSiteRow){ //*6/10/2020 KJE #259 allow adding new blank rows to Staff Lists tab
        addBlankRow(listVals); 
      }
      listVals[listValSiteRow][listSheetSiteCol-1] = thisSite;
      listValSiteRow++;
      lastSite = thisSite;
    }
    var thisProg = munisVals[i][munisSheetProgramCol-1];
    Logger.log(thisProg);
    Logger.log(listValProgRow);
    if(thisProg != lastProg){
      if(listVals.length <= listValProgRow){ //*6/10/2020 KJE #259 allow adding new blank rows to Staff Lists tab
        addBlankRow(listVals); 
      }
      listVals[listValProgRow][listSheetProgCol-1] = thisProg;
      listValProgRow++;
      lastProg = thisProg;
    }
  }
  //*6/10/2020 KJE #259 clear anything left at the bottom
  for(var i = listValSiteRow; i < listVals.length; i++){
    listVals[listValSiteRow][listSheetSiteCol-1] = "";
    listVals[listValSiteRow][listSheetSiteORCol-1] = "";
  }
  for(var i = listValProgRow; i < listVals.length; i++){
    listVals[listValProgRow][listSheetProgCol-1] = "";
    listVals[listValProgRow][listSheetProgORCol-1] = "";
    listVals[listValProgRow][listSheetPARetCol-1] = "";
    listVals[listValProgRow][listSheetPANewCol-1] = "";
  }
  //fill in overrides from previous values
  for(var i = listSheetLastHeader; i < listVals.length; i++){
    //clear old value
    listVals[i][listSheetSiteORCol-1] = "";
    listVals[i][listSheetProgORCol-1] = "";
    listVals[i][listSheetPARetCol-1] = "";
    listVals[i][listSheetPANewCol-1] = "";
    var foundSite = false;
    var foundProg = false;
    for(var k = listSheetLastHeader; k < listValsOld.length; k++){
      if(listVals[i][listSheetSiteCol-1] == listValsOld[k][listSheetSiteCol-1]){
        listVals[i][listSheetSiteORCol-1] = listValsOld[k][listSheetSiteORCol-1];
        foundSite = true;
      }
      if(listVals[i][listSheetProgCol-1] == listValsOld[k][listSheetProgCol-1]){
        listVals[i][listSheetProgORCol-1] = listValsOld[k][listSheetProgORCol-1];
        listVals[i][listSheetPARetCol-1] = listValsOld[k][listSheetPARetCol-1];
        listVals[i][listSheetPANewCol-1] = listValsOld[k][listSheetPANewCol-1];
        foundProg = true;
      }
      if(foundSite && foundProg){
        break; 
      }
    }
  }
  //set!
  if(listRange.getLastRow() < listVals.length){ //*6/10/2020 KJE #259 allow adding new blank rows to Staff Lists tab
    listSheet.getRange(1,1,listVals.length,listSheet.getMaxColumns()).setValues(listVals);
  }else{
    listRange.setValues(listVals); 
  }
}
function addBlankRow(list){ //*6/10/2020 KJE #259 allow adding new blank rows to Staff Lists tab
  var width = list[0].length;
  var newRow = [];
  for(var i = 0; i < width; i++){
    newRow.push(null); 
  }
  list.push(newRow);
}
function formatHiredFrom(){
  getUi();
  var response = ui.prompt("From which row number?","START", ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton() == ui.Button.CANCEL){return;}
  var from = response.getResponseText();
  if(isNaN(from)){ui.alert("Input must be a number. Please try again.", ui.ButtonSet.OK);return;}
  if(from <= hiredSheetLastHeader){ui.alert("Input must be greater than the header row. Please try again.",ui.ButtonSet.OK);return;}
  formatHired(from);
}
var formatHiredTimeLimit = 5.5*60*1000; //5 mins 30 seconds
function formatHired(from){
  var timer = String(new Date().valueOf()); //keep as a string to prevent converting to scientific
  getHiredSheet();
  var maxCols = hiredSheet.getMaxColumns();
  getUi();
  from = (from) ? from : hiredSheetLastHeader+1;
  for(var i = from; i <= hiredSheet.getLastRow(); i++){
    if(String(new Date().valueOf()) - timer > formatHiredTimeLimit){
      ui.alert("Timed out on row " + i + ". Start again from there.");
      return;
    }
    var rowRange = hiredSheet.getRange(i,1,1,maxCols);
    rowRange.clearFormat();
    var status = rowRange.getValues()[0][SARLib.statusCol-1];
    setBorders(hiredSheet,i,rowRange);
    if(i == hiredSheetLastHeader+1){
      hiredSheet.getRange(hiredSheetLastHeader,1,1,maxCols).setBorder(null,null,true,null,null,null) 
    }
    setStatusColor(status,hiredSheet,i,rowRange);
    var bNumRange = hiredSheet.getRange(i,bNumCol,1,1);
    if(bNumRange.getValue() == newEmpKey){
      bNumRange.setFontWeight("bold"); 
    }
  }
}
//*3/21/24 KJE rider on #s 517 and 429, make it easier to nicely format the Staff Lists tab
function formatStaffLists(){
  getListSheet();
  var maxRows = listSheet.getMaxRows();
  var maxColumns = listSheet.getMaxColumns();
  //first set whole sheet to grey
  listSheet.getRange(1,1,maxRows,maxColumns).setBorder(true,true,true,true,true,true,"grey",SpreadsheetApp.BorderStyle.SOLID).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  //Top row
  listSheet.getRange(1,1,1,maxColumns).setBorder(false,false,true,false,false,false,"black",SpreadsheetApp.BorderStyle.SOLID_THICK);
  //Columns
  var boldLines = SARLib.listSheetBoldColumns;
  for(var i = 0; i < boldLines.length; i++){
    listSheet.getRange(1,boldLines[i],maxRows,1).setBorder(null,null,null,true,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
}
function earlyReceiptCount(){
  makePleaseWait();
  var listVals = getListSheet().getDataRange().getDisplayValues();
  var earlyList = [];
  var earlyCol = SARLib.earlyReceiptsCol;
  for(var i = listSheetLastHeader; i < listVals.length; i++){
    if(listVals[i][earlyCol-1] == "Yes"){
      earlyList.push(listVals[i][listSheetSiteCol-1]); 
    }
  }
  if(earlyList.length == 0){
    getUi().alert("No sites are set to send their receipts early.");
    return;
  }
  var hiredVals = getHiredSheet().getDataRange().getDisplayValues();
  var activeVals = getActiveSheet().getDataRange().getDisplayValues();
  var both = hiredVals.concat(activeVals);
  var count = 0;
  var total = 0;
  for(var i = 0; i < both.length; i++){
    if(earlyList.indexOf(both[i][siteCol-1]) != -1){
      count++;
    }
    total++;
  }
  getUi().alert("If all " + total + " staff listed on the hired and active tabs log time this pay period, " + count + " will get their receipts early.\n\nThat leaves " + (total-count) +
    " who will get their receipts regularly.\n\nThe daily e-mail limit is 1500, including all e-mails sent by SEDbot, even non-receipts.");
  closePleaseWait();
}
function sendSiteTotals() {
  makePleaseWait();
  var hiredVals = getHiredSheet().getDataRange().getDisplayValues();
  var activeVals = getActiveSheet().getDataRange().getDisplayValues();
  var both = hiredVals.concat(activeVals);
  var totals = {};
  var sorter = [];
  for(var i = 0; i < both.length; i++){
    var site = both[i][siteCol-1];
    if(totals[site] === undefined){
      totals[site] = 0;
      sorter.push(site);
    }
    totals[site]++;
  }
  var text = "";
  sorter.sort();
  var total = 0;
  for(var i = 0; i < sorter.length; i++){
    var site = sorter[i];
    var count = totals[site];
    if(!site){
      text = text + "(no site): " + count + "\n";
    }else{
      text = text + site + ": " + count + "\n";
    }
    total = total + count;
  }
  text = text + "Total: " + total;
  GmailApp.sendEmail(getUser().getEmail(), "Site Totals", text);
  closePleaseWait();
}
function reAgeAtStart(){
  makePleaseWait();
  var sheet = SpreadsheetApp.getActiveSheet();
  if(sheet.getName() != hiredSheetName){
    getUi().alert("This option can only be used on the \"" + hiredSheetName + "\" sheet.");
    return;
  }
  var rows = getActiveRows();
  if(!rows){
    getUi().alert("Select more than just the header and try again.");
    return;
  }
  for(var i = 0; i < rows.length; i++){
    updateAgeAtStart(sheet,rows[i]);
  }
  closePleaseWait();
}
/**
 * Google periodically turns off triggers (once a year?). Every six months, send a reminder email to the PAs to reset their triggers, and send a reminder email to the Masters to reset THIS trigger.
 * 
 * This is the menu function to turn the trigger ON.
 * 
 * 3/20/2024 KJE #352 created
 * 
 * @author Kyle Egerdal
 */
function turnResetReminderOn(){
  makePleaseWait();
  ScriptApp.newTrigger("sendReminderEmails").timeBased().everyWeeks(24).onWeekDay(ScriptApp.WeekDay.MONDAY).create();
  closePleaseWait();
}
/**
 * Google periodically turns off triggers (once a year?). Every six months, send a reminder email to the PAs to reset their triggers, and send a reminder email to the Masters to reset THIS trigger.
 * 
 * This is the menu function to turn the trigger OFF.
 * 
 * 3/20/2024 KJE #352 created
 * 
 * @author Kyle Egerdal
 */
function turnResetReminderOff(){
  makePleaseWait();
  var triggers = ScriptApp.getProjectTriggers();
  for(var i = 0; i < triggers.length; i++){
    var trigger = triggers[i];
    if(trigger.getHandlerFunction() == "sendReminderEmails"){
      ScriptApp.deleteTrigger(trigger);
      closePleaseWait();
      return;
    }
  }
  closePleaseWait();
}
/**
 * Google periodically turns off triggers (once a year?). Sends a reminder email to the PAs to reset their triggers, and send a reminder email to the Masters to reset the trigger that does this.
 * 
 * This is the function that actually sends the emails. It needs a trigger to be set up by resetReminderEmail.
 * 
 * 3/20/2024 KJE #352 created
 * 
 * @author Kyle Egerdal
 */
function sendReminderEmails(){
  var localTesting = false;
  var listSheetValues = getListSheet().getDataRange().getValues();
  var masters = getMasters(listSheetValues);
  var pas = getPAs(listSheetValues);
  var developers = SARLib.errorEmailAddresses;
  if(localTesting){
    Logger.log("masters: " + masters);
    Logger.log("pas: " + pas);
    masters = developers;
    pas = developers;
  }
  //PAs
  tryTryAgain(function(){
    GmailApp.sendEmail(pas,"[IMPORTANT] Biannual Reminder: Please Turn SEDbot Off and On Again","Hello,\n\nPeriodically, SEDbot gets tired and turns itself off so it can have a break.\n\nRather than let SEDbot turn itself off without you knowing, instead, please turn SEDbot off and then on again yourself.\n\nYou can do this from within the SED. Click the \"PA\" menu at the top, then choose \"Turn off SEDbot\". After that's done, click the \"PA\" menu again and choose \"Turn on SEDbot\".\n\nIf you don't do this yourself, SEDbot will at some point turn itself off without you knowing.\n\nIf you don't use SEDbot, you can ignore this email.\n\nThank you!\nSEDbot");
  });
  //Masters
  tryTryAgain(function(){
    GmailApp.sendEmail(masters,"[IMPORTANT] Biannual Reminder: Please Turn This Reminder Off and On Again","Hello,\n\nPeriodically, SEDbot gets tired and turns itself off so it can have a break.\n\nRather than let SEDbot turn itself off without the PAs knowing, instead, we ask them to turn SEDbot off and then on again themselves. SEDbot just sent that email out now.\n\nAs the SEDbot Masters, you need to turn this reminder off and then on again, or it will stop working, too.\n\nYou can do this from within the SED. Click the \"Masters\" menu at the top, then choose \"Turn off SEDbot Reminders\". After that's done, click the \"Masters\" menu again and choose \"Turn on SEDbot Reminders\".\n\nIf you don't do this yourself, SEDbot will at some point turn these reminders off without you knowing.\n\nThank you!\nSEDbot");
  });
  //Developers
  tryTryAgain(function(){
    GmailApp.sendEmail(developers,"[IMPORTANT] Biannual Reminder: Please Turn Triggers Off and On Again","Hello,\n\nPeriodically, SEDbot gets tired and turns itself off so it can have a break.\n\nRather than let SEDbot turn itself off without you knowing, instead, you can turn SEDbot off and then on again yourself.\n\nTo do this, open each script (the SED Script and Formhandler) and run Triggers.takeDownTriggers(), then Triggers.setUpTriggers().\n\nIf you don't do this yourself, SEDbot will at some point turn the system off without you knowing.\n\nThank you!\nSEDbot");
  });

}