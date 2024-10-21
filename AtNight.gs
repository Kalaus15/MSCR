//Copyright (c) Kyle Egerdal 2018. All Rights reserved.

/**
 * Because employee paperwork folders can be created by anyone, we have to have a system for deleting them that includes all the possible "anyone"s running a script to delete them, and knowing which ones to delete.
 * 
 * So, we tag on " - DELETE" to the folder name if it's ready to be deleted, then have the PA scripts that run overnight for each PA look for folders with " - DELETE" in the name and delete them.
 * 
 * This utility checks for folder that should have been marked as ready to delete, but weren't. It assumes all folders in the folders folder that are NOT currently referenced in the SED should have been deleted. When it finds a folder that should have been marked " - DELETE" but wasn't, it marks it.
 * 
 * @author Kyle Egerdal
 */
function clearFoldersUtil(){
  var formulas = getFormulas();
  var ids = [];
  for(var i = 1; i < formulas.length; i++){
    var thisId = getIDfromCell(i,folderCol-1);
    if(thisId){
      ids.push(thisId);
    }
  }
  var folder = null;
  if(env == "PRD"){
    folder = DriveApp.getFolderById(empFolderPRDID);
  }else{
    folder = DriveApp.getFolderById(empFolderID);
  }
  var subfolders = folder.getFolders();
  var me = getUser().getEmail().toUpperCase();
  while(subfolders.hasNext()){
    var subfolder = subfolders.next();
    if(ids.indexOf(subfolder.getId()) == -1){
//      if(subfolder.getOwner().getEmail().toUpperCase() != me){
//        continue;
//      }
      var name = subfolder.getName();
      if(name.indexOf("- DELETE") == -1){
        subfolder.setName(name + " - DELETE");
      }
    }
  }
}
/**
 * Function for anyone who has created employee paperwork folders to run to delete all employee paperwork folders they own.
 * 
 * Ideally, called by PA overnight functions each night.
 * 
 * @author Kyle Egerdal
 */
function clearEmpFolders(){
  var folder = null;
  if(env == "PRD"){
    folder = DriveApp.getFolderById(empFolderPRDID);
  }else{
    folder = DriveApp.getFolderById(empFolderID);
  }
  var subfolders = folder.getFolders();
  var me = getUser().getEmail().toUpperCase();
  while(subfolders.hasNext()){
    var subfolder = subfolders.next();
    try{
      if(subfolder.getOwner().getEmail().toUpperCase() != me){
        continue;
      }
      if(subfolder.getName().indexOf("- DELETE") != -1){
        deleteMyFolder(subfolder);
      }
    }catch(e){
      if((e.message).indexOf("permission") != -1 || (e.message).indexOf("server error") != -1 || (e.message).indexOf("Service error") != -1 || (e.message).indexOf("Limit Exceeded") != -1){
        continue; //try again tomorrow
      }else{
        throw e;
      }
    }
  }
}
/**
 * Wrapper for the atNight function to notify developers of any errors.
 * 
 * @author Kyle Egerdal
 */
function atNightTrigger(){
  try{
    atNight();
  }catch(e){
    GmailApp.sendEmail(SARLib.getErrorEmails(env), "atNight() Error", e + "\n" + e["stack"]);
  }
}
var daysToKeepArchives = 30;
var bkDateEx = /\d{2}\/\d{2}\/\d{2}/; //Used to detect backup files by the way their date is formatted
/**
 * Does various functions related to the SED overnight, including:
 * - Moving staff between tabs of the spreadsheet by calling moveViaStatus. Gets a lock to do this.
 * - Making a backup of the SED
 * - Deleting any SED backups older than daysToKeepArchives days
 * - Deletes any employee paperwork folders created by the system that are tagged for deletion by calling clearEmpFolders
 * - Archives any cover sheets older than 30 days by calling archiveCoverSheets
 * 
 * Not to be run at the same time as PA overnight functions to avoid desync issues. The locking system should handle this anyway, but why run simultaneously when you know the locks will prevent one or the other from running anyway?
 * 
 * @param {boolean} [secondTry] Whether the function is trying again after not being able to get a lock. The default is false; i.e., assume unless told otherwise that this is the first run of this function.
 * @author Kyle Egerdal
 */
function atNight(secondTry){
  if(secondTry !== true){ //*6/7/19 KJE !secondTry is always false when this function is triggered
    SARLib.startTriggerTracker(env, "SED-atNight");
  }
  if((new Date()).getDay() == 0){return;} //don't run on Sundays -- can't get a lock from FormHandler.!EmpHandler pushSARToVortex
  getHiredSheet();
  var listSheetVals = getListSheet().getDataRange().getValues();
  SARLib.massUnlock(env,"@madison.k12.wi.us");
  //#485 YMC 9/12/23 lock atNight and unlock it when it's done to prevent another atNight running the same time
 if(!SARLib.SARWriteLock("atNight",env,hiredSheet,getMasters(listSheetVals))){
   cleanUpTriggers("atNightTrigger"); //HelperFxns.gs
   try{
     ScriptApp.newTrigger("atNightTrigger")
     .timeBased()
     .after(900000) //try again in 15 minutes
     .create();
   }catch(e){
     if((e.message).indexOf("server error") != -1){
       Utilities.sleep(2000);
       atNight(true);
     }else{
       throw e;
     }
   }
   return;
 }
  var error;
  try{
    getSSByID();
    moveViaStatus(listSheetVals);
  }catch(e){
    error = e;
  }
 SARLib.SARWriteUnlock(env,hiredSheet);
  if(error){
    GmailApp.sendEmail(SARLib.getErrorEmails(env), env + ": atNight() Error", error + error["stack"]);
  }
  if(env == "PRD"){
    var SED = getSSByID();
    //make backup
    var bk = DriveApp.getFileById(SED.copy(SED.getName() + " - Backup " + Utilities.formatDate(new Date(), tz, "MM/dd/yy")).getId());
    var bkF = SARLib.getBackupFolder();
    bkF.addFile(bk);
    DriveApp.removeFile(bk);
    //delete old backups
    var deleteAfter = Utilities.formatDate(new Date(new Date().valueOf() - daySecs*daysToKeepArchives), tz, "MM/dd/yy");
    var deleteAfterParts = deleteAfter.split("/");
    var it = bkF.getFiles();
    while(it.hasNext()){
      var file = it.next();
      var name = file.getName();
      if(name.indexOf("SED - PRD - Backup") == -1){continue};
      var date = name.match(bkDateEx)[0];
      var dateParts = date.split("/");
      if(dateParts[2] > deleteAfterParts[2]){continue;}    //after this year, skip
      if(dateParts[2] == deleteAfterParts[2]){              //this year, more evaluation needed
        if(dateParts[0] > deleteAfterParts[0]){continue;}    //after this month, skip
        if(dateParts[0] == deleteAfterParts[0]){              //this month, more evaluation needed
          if(dateParts[1] > deleteAfterParts[1]){continue;}    //after today, skip
          //today or earlier, delete
        }                                            //before this month, delete
      }                                            //before this year, delete
      file.setTrashed(true);
    }
  }
  clearEmpFolders();
  archiveCoverSheets();
  SARLib.documentTriggerTracking(env, "SED-atNight");
}
/**
 * Given a row, checks to see if that row has an end date for the contract. If so, and the staff didn't quit or get terminated (as represented by the presence of a quit date), checks to see if the end date is more recent than the current most recent employment date for the staff, if stored. If so, or if no date is stored, updates the staff's most recent employment date to be the row's quit date.
 * 
 * @param {object} row The row in question.
 * @param {object} needMRED The needMRED object whose keys are MREDIds and whose values are the most recent employment dates of the employees with those IDs.
 * @param {String} [MREDId] The MREDId for the row in question, as returned by getMREDId, if known. If not known, calculated. 
 * @author Kyle Egerdal
 */
function checkMRED(row,needMRED,MREDId){
  if(!MREDId){
    MREDId = getMREDId(row);
  }
  if(needMRED[MREDId] !== undefined){
    var curMRED = needMRED[MREDId];
    var newMRED = new Date(row[endDateCol-1]); //*4/1/20 KJE #229 wasn't using dates encoutered later in the script
    if(newMRED && !row[quitDateCol-1]){
      if(curMRED){
        if(newMRED > curMRED){
          needMRED[MREDId] = newMRED.valueOf();
        }
      }else{
        needMRED[MREDId] = newMRED.valueOf();
      }
    }
  }
}
/**
 * Checks to see if we need the birthday for an employee on the "Hired" tab. If so, sets it.
 * 
 * Uses the MREDId as the staff identifier.
 * 
 * 3/21/24 KJE #404 created
 * 
 * @param {object} row The staff's row on the Active or Archive tab.
 * @param {object} needBirthday The needBirthday object used to store staff who need their birthdays updated.
 * @param {String} [MREDId] The MREDId for the row, calculated if not given.
 * @author Kyle Egerdal
 */
function checkBirthday(row,needBirthday,MREDId){
  if(!MREDId){
    MREDId = getMREDId(row);
  }
  if(needBirthday[MREDId] == ""){
    var birthday = row[birthdayCol-1];
    if(birthday == "18+"){
      needBirthday[MREDId] = birthday;
    }else{
      needBirthday[MREDId] = new Date(row[birthdayCol-1]);
    }
  }
}
/**
 * Calculates a most recent employment date ID for an employee, which is just the all-uppercase combination of their first name, last name, and email address. This ID needs to be unique enough to distinguish individual staff, but generic enough it can identify duplicates without a B# (since there is no guarantee the person inputting staff on the SED will know whether that employee worked for MSCR in the past or not).
 * 
 * If an employee has changed their email address since the last time we employed them, this won't recognize them. This is an acceptable failure rate, as it is more successful than trying to idenfying previous staff by B# (see inputting issue above).
 * 
 * @param {object} row Thw row for which to generate an MREDId.
 * @return {String} The MREDId for the given row.
 * @author Kyle Egerdal
 */
function getMREDId(row){
  return (row[firstNameCol-1] + row[lastNameCol-1] + row[emailCol-1]).toUpperCase();
}
var MREDNotFound = "SEDbot: Not Found";
/**
 * Moves staff between tabs appropriately. That is,
 * - Moves staff on the "Active" tab whose contract expired three or more weeks ago (or they were terminated three or more weeks ago) to the "Archive" tab. We use three or more weeks because this ensures their information is on the "Active" tab when their CSV is being generated, and thus CSV generation does not have to use the "Archive" tab on the SED, which would take much longer. This also allows staff to use manual entry to input time they may have worked at the beginning of a pay period, even if it's now the end of that pay period and their contract is inactive.
 * - Moves staff on the "Hired" tab who are cleared to work to the "Active" tab. If the PA who does their paperwork is different than the PA who uploads their time, changes the PA they have listed from the former to the latter.
 * - Deletes staff on the "Hired" tab whose status is set to "Delete."
 * 
 * All the while:
 * - Updates birthdays and most recent employment dates for staff who need it, and for whom those dates can be found.
 * - Fixes duplicate TCIDs, if it finds any and they aren't yet in use.
 * - Notifies developers of duplicate TCIDs for cleanup, if it finds any and they ARE in use.
 * - If an employee logged time before they were cleared, and they are now being cleared, notifies their PA.
 * 
 * And finally:
 * - Notifies PAs of staff they cleared that were successfully moved to the "Active" tab.
 * - Notifies PAs of staff they were assigned.
 * - Reformats the "Active" sheet in case copying staff to it from the "Hired" tab ruined the formatting.
 * - Sorts the "Active" and "Archive" tabs by last name, then first name.
 * - Various other minor cleanups and touch-ups such as deleting extra rows and setting permissions.
 * 
 * 11/9/2021 MDH #336 replaced functionality
 * 2/28/22 MDH #385 fix sorting
 * 
 * @param {object} listSheetVals The values in the list sheet.
 * @author Kyle Egerdal
 */
function moveViaStatus(listSheetVals){
  getSSByID();
  var hiredSheet = getHiredSheet();
  var hiredRange = hiredSheet.getDataRange(); //used by getVals
  var hiredVals = hiredRange.getDisplayValues(); //*5/6/20 KJE #248 explicitly define hiredVals to prevent conflation
  getFormulas();
  //first get hires that need a MRED
  var needMRED = {};
  var needBirthday = {}; //*3/21/24 KJE #404 fill in birthdays, too
  var splitChecker = {}; //*5/13/20 KJE #149 notify of split agreements
  for(var i = hiredSheetLastHeader; i < hiredVals.length; i++){
    var row = hiredVals[i];
    var MREDId = getMREDId(row);
    if(!row[recEmpCol-1]){
      var row = hiredVals[i];
      needMRED[MREDId] = "";
    }
    //*3/21/24 KJE #404
    var birthday = row[birthdayCol-1];
    if(!birthday || birthday == "18+"){
      needBirthday[MREDId] = "";
    }
    var bNum = row[bNumCol-1]; //*5/13/20 KJE #149 notify of split agreements
    if(bNum && bNum != "NEW"){ //*5/13/20 KJE #149 notify of split agreements
      if(splitChecker[bNum] === undefined){
        splitChecker[bNum] = {};
      }
      splitChecker[bNum][row[idCol-1]] = row[jobCodeCol-1];
    }
  }
  //then go through archive to see if any are in there
  var archiveVals = getArchiveSheet().getDataRange().getDisplayValues();
  for(var i = SARLib.archiveLastHeader; i < archiveVals.length; i++){
    var row = archiveVals[i];
    var MREDId = getMREDId(row);
    checkMRED(row,needMRED,MREDId);
    checkBirthday(row,needBirthday,MREDId); //*3/21/24 KJE #404
  }
  //then move from active to archive. since hired will add to active we'll go through fewer rows if we go through active first.
  //all the while, track MRED
  var activeVals = getActiveSheet().getDataRange().getValues();
  var archiveBuffer = 604800000*3; //three weeks in milliseconds
  var archiveDeadline = ((new Date()).valueOf() - archiveBuffer);
  var duplChecker = {};

  var deleted = 0;
  for(var i = activeSheetLastHeader; i < activeVals.length; i++){
    var row = activeVals[i];
    var MREDId = getMREDId(row);
    checkMRED(row,needMRED,MREDId);
    checkBirthday(row,needBirthday,MREDId); //*3/21/24 KJE #404
    var id = row[idCol-1];
    if(duplChecker[id]){
      SARLib.sendAdminEmail(env, "Duplicate TCID IN USE NOW! ID" + id);
    }else{
      duplChecker[id] = 1; //anything that evaluates to true
    }
    var bNum = row[bNumCol-1];  //*5/13/20 KJE #149 notify of split agreements
    if(splitChecker[bNum] !== undefined){ //*5/13/20 KJE #149 notify of split agreements
      splitChecker [bNum][row[idCol-1]] = row[jobCodeCol-1];
    }
    var removeDate = row[quitDateCol-1];
    if(!removeDate){
      removeDate = row[endDateCol-1];
    }
    var status = row[statusCol-1];
    var rValue = new Date(removeDate).valueOf();
    //*5/6/20 KJE #248 don't archive if rvalue is blank
    if(rValue && rValue < archiveDeadline){// || status == quit){ *6/3/19 KJE don't move quit staff until 3 week buffer -- need info for receipts
      if(status != quit){
        row[dateAddedCol-1] = new Date();
      }
      archiveSheet.appendRow(row);
      activeSheet.deleteRow(i+1-deleted); //vals is 0-indexed, sheets is 1-indexed
      deleted++;
    }
  }
  //finally, move from hired to active + archive and set MRED
  var deleted = 0;
  var activeBefore = activeSheet.getMaxRows();
  var uEmail = getUser().getEmail().toUpperCase();
  var today = new Date();
  if(!testing){
    today.setDate(today.getDate()-1);
  }
  var todayFmtd = Utilities.formatDate(today, tz, "M/d/YYYY"); //*5/13/2020 KJE #241 consolidate notification emails
  var paNotifObj = {};
  for(var i = hiredSheetLastHeader; i < hiredVals.length; i++){
    var row = hiredVals[i];
    var MREDId = getMREDId(row);
    if(!hiredVals[i][recEmpCol-1]){
      var newMRED = needMRED[MREDId];
      if(newMRED !== undefined){
        if(newMRED){ //exists and isn't empty
          var range = hiredSheet.getRange(i+1-deleted,recEmpCol);
          range.setValue(new Date(newMRED));
          range.setNumberFormat("M/d/yyyy");
          updateAgeAtStart(hiredSheet,i+1-deleted);
        }else{
          var tempRange = hiredSheet.getRange(i+1-deleted,recEmpCol);
          tempRange.clearDataValidations(); //*4/24/23 KJE #490 handle putting text in a date-formatted cell, which throws an error on next batch upload
          tempRange.setValue(MREDNotFound);
        }
      }
    }
    //*3/21/24 KJE #404
    var newBirthday = needBirthday[MREDId];
    if(newBirthday){
      if(row[birthdayCol-1] != newBirthday){
        hiredSheet.getRange(i+1-deleted,birthdayCol).setValue(newBirthday);
        updateAgeAtStart(hiredSheet,i+1-deleted);
      }
    }
    var id = row[idCol-1];
    var newId = null;
    if(id && id != " " && duplChecker[id]){
      newId = "#" + getNextEntryID();
      row[idCol-1] = newId;
      duplChecker[newId] = 1;  //anything that evaluates to true
      if(id && id != " "){ //*3/11/23 KJE hotfix: don't notify for nullish ids
        SARLib.sendAdminEmail(env, "Fixing duplicate TCID. Old ID: " + id + " New ID: " + newId);
      }
    }else{
      duplChecker[id] = 1; //anything that evaluates to true
    }
    var status = row[statusCol-1];
    if(status){
      if((status == cleared) || (status == replaced)){ //*11/9/2021 MDH #336 replaced functionality
        //remove spaces from first and last name
        row[firstNameCol-1] = String(row[firstNameCol-1]).trim();
        row[lastNameCol-1] = String(row[lastNameCol-1]).trim();
        row[dateAddedCol-1] = new Date();
        var firstDateLogged = row[firstDateLoggedCol-1];
        row.splice(lastKeepCol); //don't keep info on paperwork

        //*9/24/2019 KJE #160 notify PAs when their staff are cleared
        //*5/13/2020 KJE #241 consolidate notifications
        var oldPA = row[PACol-1];
        if(oldPA){
          if(paNotifObj[oldPA] === undefined){  //*5/13/2020 KJE #241 consolidate notification emails
            paNotifObj[oldPA] = {};
          }
          if(paNotifObj[oldPA]["cleared"] === undefined){
            paNotifObj[oldPA]["cleared"] = [row.slice()];
          }else{
            paNotifObj[oldPA]["cleared"].push(row.slice());
          }
        }
        //set PA to PA who does timesheet
        var PA;
        PA = SARLib.getPA(row[programCol-1],false,listSheetVals);
        if(PA){
          row[PACol-1] = PA;
        }else{
          PA = row[PACol-1];
        }
        activeSheet.appendRow(row);
        if(hiredSheet.getMaxRows() > hiredSheetLastHeader+1){
          hiredSheet.deleteRow(i + 1 - deleted);
        }
        if(firstDateLogged){// &&
          var firstDateAsDate = new Date(firstDateLogged); //*11/22/20 KJE let's try this again
          //*6/16/21 KJE add diff > 0 to handle same day time logging
          var diff = (Math.floor(today.getTime()/86400000) - Math.floor(firstDateAsDate.getTime()/86400000));
          if(diff > 0 && !(diff <= 7 &&
            (today.getDay() - firstDateAsDate.getDay() > 0))){
              //if firstDatelogged was less than a week ago and it was the same week (ie there is not a Saturday at midnight in between)
              //we know it's the same pay period. So don't send e-mail.
              sendPAUploadEmail(PA,row,firstDateLogged);
            }
        }
        deleteFolder(row,uEmail);
        deleted++;
      }else if(status == quit){
        row[dateAddedCol-1] = new Date();
        row.splice(lastKeepCol); //don't keep info on paperwork
        archiveSheet.appendRow(row);
        hiredSheet.deleteRow(i + 1 - deleted);
        deleteFolder(row,uEmail);
        deleted++;
      }else if(status == deleteMe){
        hiredSheet.deleteRow(i + 1 - deleted);
        deleteFolder(row,uEmail);
        deleted++;
      }else if(newId){
        hiredSheet.getRange(i + 1 - deleted,idCol).setValue(newId);
      }
    }else{
      if(newId){
        hiredSheet.getRange(i + 1 - deleted,idCol).setValue(newId);
      }
      if(row[dateAddedCol-1] == todayFmtd){  //*5/13/2020 KJE #241 consolidate notification emails
        var PA = row[PACol-1];
        checkNewStaffInfo(row, PA);
        if(PA){
          if(paNotifObj[PA] === undefined){
            paNotifObj[PA] = {};
          }
          if(paNotifObj[PA]["new"] === undefined){
            paNotifObj[PA]["new"] = [row.slice()];
          }else{
            paNotifObj[PA]["new"].push(row.slice());
          }
        }
      }
    }
  }
  //*5/13/2020 KJE #241 consolidate notification emails
  //*9/24/2019 KJE #160 notify PAs of cleared staff
  for(pa in paNotifObj){
    var text = "Hello,\n\n";
    if(paNotifObj[pa]["new"] !== undefined){
      var list = paNotifObj[pa]["new"];
      text = text + "You were assigned the following staff today:\nID# - Name - Notes";
      for(var i = 0; i < list.length; i++){
        var row = list[i];
        var TCID = row[idCol-1];
        var bNum = row[bNumCol-1];
        var job = row[jobCodeCol-1];
        var ids = splitChecker[bNum];
        var splitId = null;
        for(thisId in ids){
          if(thisId == TCID){
            continue;
          }
          if(ids[thisId] == job){
            splitId = thisId;
            break;
          }
        }
        text = text + "\n" + row[idCol-1] + " - " + row[firstNameCol-1] + " " + row[lastNameCol-1] + (splitId ? " - Split with ID" + splitId : "");
      }
      text = text + "\n\n";
    }
    if(paNotifObj[pa]["cleared"] !== undefined){
      var list = paNotifObj[pa]["cleared"];
      text = text + "The following staff of yours were cleared and moved to the \"Active\" tab in the SED just now:\nID# - Name";
      for(var i = 0; i < list.length; i++){
        var row = list[i];
        text = text + "\n" + row[idCol-1] + " - " + row[firstNameCol-1] + " " + row[lastNameCol-1];
      }
    }
    text = text + "\n\nThank you,\n" + SARLib.robotName;
    GmailApp.sendEmail((testing ? SARLib.getErrorEmails(env) : pa), "SED Staff Updates - " + todayFmtd, text);
  }
  //delete extra rows
  var hlr = sheet.getLastRow();
  var hmr = sheet.getMaxRows();
  if((hmr > hiredSheetLastHeader + 1) && (hmr > hlr)){
    sheet.deleteRows(hlr+1,hmr-hlr);
  }
  //set row permissions
  tryTryAgain(function(){  //*11/14/20 KJE tryTryAgain
    setRowPermissions();
  });
  //set active sheet validation, and active + archive formatting
  var amr = activeSheet.getMaxRows();
  var rows = amr-activeSheetLastHeader;
  if(rows > 1){
    //clear formatting
    activeSheet.getRange(activeBefore+1,1,amr,lastKeepCol).clearFormat();
    //clear data validations, but...
    var asmc = activeSheet.getMaxColumns();
    activeSheet.getRange((activeSheetLastHeader+1),1,rows,(asmc-1)).clearDataValidations();
    //re-set status
    activeSheet.getRange(activeSheetLastHeader+1,statusCol,rows,1).setDataValidation(SpreadsheetApp.newDataValidation()
                                                                                     .requireValueInList([cleared,quit,replaced], true)
                                                                                     //*11/9/2021 MDH #336 add replaced 
                                                                                     .setAllowInvalid(false).build());
    //re-set supervisor and payroll approver
    getListSheet();
    var listSheetLastRow = listSheet.getLastRow();
    var validCol = getANotation(listSheetSDCol);
    var lastvalidCol = getANotation(listSheetPACol);
    var validRange = listSheet.getRange(validCol+(listSheetLastHeader+1)+":"+lastvalidCol+listSheetLastRow);
    var supValidation = SpreadsheetApp.newDataValidation().requireValueInRange(validRange, true).setAllowInvalid(false).build();
    activeSheet.getRange(activeSheetLastHeader+1,supCol,rows,1).setDataValidation(supValidation).setHorizontalAlignment("left");
    activeSheet.getRange(activeSheetLastHeader+1,payrollApvsCol,rows,1).setDataValidation(supValidation).setHorizontalAlignment("left");
    activeSheet.getRange(activeSheetLastHeader+1,SARLib.remindCol,rows,1).setDataValidation(SpreadsheetApp.newDataValidation()
                                                                                            .requireValueInList(SARLib.remindList)
                                                                                            .setAllowInvalid(false).build());
    activeSheet.getRange(activeSheetLastHeader+1,1,rows,asmc).sort([lastNameCol,firstNameCol]);
  }
  archiveSheet.getRange(activeSheetLastHeader+1,1,archiveSheet.getMaxRows()-activeSheetLastHeader,archiveSheet.getMaxColumns()).clearFormat().sort([lastNameCol,firstNameCol]); //*4/1/2020 KJE #185 sort alphabetically *2/28/22 MDH #385 fix sorting
}

function checkNewStaffInfo(row, PA){
  var id = row[idCol-1]; //*6/10/2020 KJE #258 "id" was missing?
  var firstName = row[firstNameCol-1];
  var lastName = row[lastNameCol-1];
  var emailAddress = row[emailCol-1];
  var startDate = row[startDateCol-1];
  if(!(firstName && lastName && emailAddress && startDate)){
    var sd = row[directorCol-1];
    var sdName = "";
    if(sd){
      sdName = sd.substring(0,sd.indexOf("<")-1);
    }
    var paName = PA ? PA.substring(0,PA.indexOf("<")-1) : "";
    var subject = robotName + ": ID" + id + " is missing information";
    var message = null;
    if(paName){
      if(sd){
        message = "Hello,<br/><br/>This e-mail is to notify you that " + sdName + "'s employee, ID" + id + ", is assigned to " +
          paName + ". However, they are missing the following information:";
      }else {
        message = "Hello,<br/><br/>This e-mail is to notify you that the employee with ID" + id + " is assigned to you." +
          "However, that employee is missing the following information:";
      }
    }else{
      return;
    }
    if(!firstName){message = message + "<br/>- First name";}
    if(!lastName){message = message + "<br/>- Last name";}
    if(!emailAddress){message = message + "<br/>- E-mail address";}
    if(!startDate){message = message + "<br/>- Start Date";}
    message = message + "<br/><br/>";
    if(sd){
      message = message + sdName.substring(0,sdName.indexOf(" ")) + ", please complete this employee's entry so " + paName.substring(0,paName.indexOf(" ")) + " can start the paperwork process."
    }else{
      message = message + "I would have notified this employee's site director, but there isn't one listed. Sorry!";
    }
    message = message + "<br/><br/>Thanks,<br/>" + robotName;
    message = message + "<br/><br/><a href=\"" + getSSByID().getUrl() + "\">Link to SED</a>";
    if(testing){
      if(sd){
        sd = SARLib.getErrorEmails(env);
      }
      PA = SARLib.getErrorEmails(env);
    }
    GmailApp.sendEmail(sd ? sd : PA,subject,"",{
      htmlBody: message,
      cc: sd ? PA : ""
    });
  }
}

function deleteFolder(row,uEmail){
  var rowFormulas = formulas[row];
  if(!rowFormulas){return;}
  var folderUrl = getUrl(rowFormulas[folderCol-1]);
  if(folderUrl){
    var folder = DriveApp.getFolderById(getIDfromUrl(folderUrl));
    deleteThisFolder(folder,uEmail);
  }
}
function deleteThisFolder(folder,uEmail){
  if(!uEmail){
    uEmail = getUser().getEmail().toUpperCase();
  }
  if(folder.getOwner().getEmail().toUpperCase() == uEmail){
    deleteMyFolder(folder);
  }else{
    var name = folder.getName();
    if(name.indexOf("- DELETE") == -1){
      folder.setName(name + " - DELETE");
    }
  }
}
function deleteMyFolder(folder){
  updateDrivePermissions([],folder,true,true); //remove any read-only access for specific users
  updateDrivePermissions([],folder,true,false); //remove any edit access for specific users
  folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); //remove any access for non-specific users
  folder.setTrashed(true); //toss it
}
function sendPAUploadEmail(PA,row,firstDateLogged){
  var fName = row[firstNameCol-1];
  var lName = row[lastNameCol-1];
  var id = row[idCol-1];
  //var date = Utilities.formatDate(firstDateLogged, tz, "MM/dd/YY"); //*6/18/2020 KJE #262 after #248 firstDateLogged is already the string date
  var html = newText("","Hello,",2);
  //*6/18/2020 KJE #262 after #248 firstDateLogged is already the string date
  html = newText(html,fName + " " + lName + " (ID" + id + ") was cleared today but began logging time as early as " + firstDateLogged + ".",2);
  html = newText(html,"If " + firstDateLogged + " is not within the current pay period, check your CSV e-mails to see if " + fName + " is ever listed with \"Unpaid time.\"" +
                 " If so, you'll need to pull their time from the <a href=\"" + magicFormUrl + "\">Magic Form,</a> once per unpaid pay period, " +
                 "and upload it to MUNIS.",2)
  html = newText(html,"Thank you,",1);
  html = newText(html,robotName);
  SARLib.sendHTMLEmail(testing ? SARLib.getErrorEmails(env) : PA, "Possible missed time for " + fName + " " + lName + " (ID" + id + ")" , html);
}
