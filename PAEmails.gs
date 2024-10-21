//Copyright (c) Kyle Egerdal 2018. All Rights reserved.

var lastRunProp = "lastRunProp_"; //add function name
var minutesFunctionList = "checkEmail"; //comma-delimited
var everyMinutes = 15; //must be 1,5,10,15, or 30
var nightlyFunctionList = "nightlyChecks";//comma-delimited
var nightlyHour = 5;
var robotName = "SEDbot";
var labelName = robotName;

function nightlyChecks(){
  var key = getUser().getEmail() + "/nightlyChecks";
  if(!SARLib.SARReadLock(key,env)){return;}
  var error;
  try{//catch any errors so we can PAunlock before throwing
    getSSByID();
    newCheckForFiles(); //first check for files
    toPA(); //then send completed employees to PA
    createReminders(); //then create reminders
    clearEmpFolders(); //finally, delete any old folders
  }catch(e){
    error = e;
  }
  SARLib.SARReadUnlock(key,env);
  if(error){
    GmailApp.sendEmail(SARLib.getErrorEmails(env), "nightlyStuffInOrder() Error", error + "\n" + error["stack"]); 
  }
}

function setUpMyTriggers(){
  makePleaseWait();
  var minutesSplit = minutesFunctionList.split(",");
  var nightlySplit = nightlyFunctionList.split(",");
  var triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
  for(var i = 0; i < triggers.length; i++) {
    var handlerFunction = triggers[i].getHandlerFunction();
    for(var j = 0; j < minutesSplit.length; j++){
      if(handlerFunction == minutesSplit[j]){ //already have this trigger
        minutesSplit.splice(j,1); //remove it from the list of triggers to add
        continue; //can only match one trigger at a time
      }
    }
    for(var j = 0; j < nightlySplit.length; j++){
      if(handlerFunction == nightlySplit[j]){
        nightlySplit.splice(j,1);
        continue;
      }
    }
  }
  if(minutesSplit.length == 0 && nightlySplit.length == 0){
    SpreadsheetApp.getUi().alert("Your triggers are already set up!");
    return;
  }
  for(var i = 0; i < minutesSplit.length; i++){
    ScriptApp.newTrigger(minutesSplit[i]).timeBased().everyMinutes(everyMinutes).create();
  }
  for(var i = 0; i < nightlySplit.length; i++){
    ScriptApp.newTrigger(nightlySplit[i]).timeBased().atHour(nightlyHour).everyDays(1).create();
  }
  closePleaseWait();
}
function deleteMyTriggers(){
  makePleaseWait();
  var triggers = ScriptApp.getProjectTriggers(); //*2/26/2019 KJE #219 change from getUserTriggers b/c of new runtime
  for(var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    if(minutesFunctionList.indexOf(trigger.getHandlerFunction()) != -1){ //just delete specified triggers, o/w admin functions like myOnEdit will go too
      ScriptApp.deleteTrigger(triggers[i]);
    }
    if(nightlyFunctionList.indexOf(trigger.getHandlerFunction()) != -1){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  closePleaseWait();
}
function getCheckEmailKey(){ //*5/1/19 KJE attempt to prevent "readProp" not defined errors
  var myEmail = getUser().getEmail();
  var key = myEmail + "/checkEmail";
  return key;
}
var idNumEx = /ID#\d+/;
/**
 * 2/2/22 MDH hotfix add tryTryAgain
 */
function checkEmail(){
  var error;
  try{//catch any errors so we can PAunlock before throwing
    var myEmail = getUser().getEmail();
    //needed variables
    myEmail = myEmail.toUpperCase();
    var since = getLastRun("checkEmail");
    if(!since){ //first run
      since = String(((new Date()).valueOf())-(everyMinutes*60*1000)).substring(0,10); //start with everyMinutes ago
      //Gmail search doesn't use MS, cut off last 3 digits
    }
    
    //get threads
    if(testing){
      var threads = GmailApp.search("in:inbox");
    }else{
      var threads = tryTryAgain(function(){ //*4/1/23 KJE hotfix tryTryAgain
        return GmailApp.search("in:inbox after:" + since);
      },null,10);
    }
    if(threads.length == 0){ //no emails since last run
      setLastRun("checkEmail");
      return;
    }
    if(!SARLib.PAReadLock(getCheckEmailKey(),env)){return;}  //*5/1/19 KJE attempt to prevent "readProp" not defined errors
    //setLastRun //*5/6/20 KJE #253 ignore not allowed error, but move setLastRun before error catch so script does the same emails again
    
    //get spreadsheet stuff
    try{ //*7/16/20 KJE #270 try again later on service error
      getSSByID();
      getVals();
      getFormulas();
    }catch(e){
      if((e.message).indexOf("Service") != -1){
        SARLib.PAReadUnlock(getCheckEmailKey(),env); //*6/19/22 YMC #399 FixIt#401: unlock first 
        //before returning caused by service unavailable error
        return; 
      }else{
        throw e; 
      }
    }
    var label = getLookedAtLabel();
    var listVals = tryTryAgain(function() { 
      return getListSheet().getDataRange().getValues();
      }); // * 2/2/22 MDH hotfix add tryTryAgain
    var ignoreFrom = getAdmins(listVals);
    ignoreFrom = ignoreFrom.concat(getPAs(listVals));
    ignoreFrom = ignoreFrom.concat(getProcessors(listVals));
    ignoreFrom = ignoreFrom.concat(getSDs(listVals));
    ignoreFrom = unique(ignoreFrom);
    ignoreFrom = ignoreFrom.toString().toUpperCase();
    
    //do the thing
    var foundOne = false;
    for(var i = 0; i < threads.length; i++){
      var thread = threads[i];
      var messages = thread.getMessages();
      for(var j = 0; j < messages.length; j++){
        try{
          var message = messages[j];
          if(message.getCc().toUpperCase().indexOf(myEmail) != -1){ //*10/22/2019 KJE #180 ignore e-mails user is CC'd on
            continue; 
          }
          var from = message.getFrom().toUpperCase();
          var bracket = from.indexOf("<");
          if(bracket != -1){
            from = from.substring(bracket+1,from.length-1); 
          }
          if(from == myEmail){
            continue; //ignore e-mails from self
          }
          var date = String(message.getDate().getTime()).substring(0,10); //Gmail search doesn't use MS, cut off last 3 digits
          if(!testing && date < since){
            continue; //thread was updated after since but message was before then. try next message
          }
          var lineNumber = null;
          var subject = message.getSubject().replace(" ","");
          if(ignoreFrom.indexOf(from) != -1){
            if(idNumEx.test(subject)){
              lineNumber = getLineNumber(null,(subject.match(idNumEx)[0]).match(/#\d+/)[0]);
            }else{
              continue; 
            }
          }else{
            lineNumber = getLineNumber(from);
          }
          if(lineNumber != -1){ //on SAR
            try{
              processEmail(lineNumber,subject,message,message.getAttachments(),myEmail.toLowerCase());
            }catch(e){
              if((e.message).indexOf("Service") != -1){
                GmailApp.sendEmail(SARLib.getErrorEmails(env), "checkEmail() error", e.message + "\n" + e["stack"]); 
              }
            }
            thread.addLabel(label);
          }
        }catch(e){
          if(e.message == "Message no longer valid."){
            continue;
          }else{
            throw e; 
          }
        }
      }
    }
  setLastRun("checkEmail"); //*5/6/20 KJE #253 ignore not allowed error, but move setLastRun before error catch so script does the same emails again
  }catch(e){
    error = e;
  }
  SARLib.PAReadUnlock(getCheckEmailKey(),env); //*5/1/19 KJE attempt to prevent "readProp" not defined errors
  if(error){
    //*5/6/20 KJE #253 ignore not allowed error, but move setLastRun before error catch so script does the same emails again
    //*8/27/20 KJE #279 ignore data storage error (do the same emails again next time)
    if((error.message).indexOf("not allowed") != -1 || (error.message).indexOf("Data storage error") != -1){
      return;
    }
    GmailApp.sendEmail(SARLib.getErrorEmails(env), "checkEmail() Error", error + error["stack"]);
  }
}
function getLookedAtLabel(){
  var label = GmailApp.getUserLabelByName(labelName);
  if(!label){
    label = GmailApp.createLabel(labelName); 
  }
  return label;
}
//assumes: vals
function getLineNumber(from,idNum){
  if(!from && !idNum){
    return -1; 
  }
  for(var i = 0; i < vals.length; i++){
    if(from){
      var thisEmail = vals[i][emailCol-1];
      try{
        if(thisEmail && from.indexOf(String(thisEmail).toUpperCase()) != -1){
          return i; 
        }
      }catch(e){ //*4/24/19 KJE if vals is corrupted, thisEmail can be a date
        if((e.message).indexOf("Cannot find function toUpperCase") != -1){
          //do nothing 
        }else{
          throw e; 
        }
      }
    }else if(idNum){
      if(vals[i][idCol-1] == idNum){
        return i; 
      }
    }
  }
  return -1;
}
//assumes: vals
function processEmail(lineNumber,subject,message,attachments,myEmail){
  var row = vals[lineNumber];
  if(attachments.length == 0){
    //*3/22/24 KJE #490 previously, bNum would be grabbed from html if there was a 6-digit color code (such as #666666 in Julia's signature). Now, only look for a bNum for new hires, and use the plain text body.
    var returning = subject.includes("RETURNING");
    var bNum = "";
    var body = message.getPlainBody();
    if(!returning){
      bNum = body.match(/\d{6}/);
      if(bNum){
        bNum = bNum[0];
      }
    }
    //*4/26/19 KJE handle ,cleared and .cleared
    //*3/22/24 KJE #490 allow clearing without bNum for returning hires
    if(subject.indexOf("Completed Paperwork") != -1 && body.replace(",",".").toUpperCase().indexOf(".CLEARED") != -1 && (returning || bNum)){
      clearEmployee(message,row,bNum,lineNumber,myEmail);
    }
    return;
  }
  var firstName = row[firstNameCol-1];
  var lastName = row[lastNameCol-1];
  var id = row[idCol-1];
  var processed = [],found = [];
  var folder = getFolder(lineNumber);
  for(var i = 0; i < attachments.length; i++){
    var attachment = attachments[i];
    var attachmentBlob = attachment.copyBlob();
    var file = DriveApp.createFile(attachmentBlob);
    file.moveTo(folder); //*3/22/24 KJE rider on #490, use not-deprecated code
    var name = file.getName();
    var recognizedAs = newMatchFile(firstName,lastName,lineNumber,file,found);
    if(recognizedAs == -1){
      //file.setTrashed(true); //*5/23/2019 KJE move to folder anyway
      processed.push(name + ": Wasn't recognized, but moved to the employee's folder anyway.");
    }else{
      processed.push(name + ": Filed as the employee's " + recognizedAs + "."); 
    }
  }
  statusEmail(message,firstName,lastName,folder,id,processed,myEmail)
}
//assumes: sheet,range
function clearEmployee(message,row,bNum,lineNumber,myEmail){
  /*
  // 10/8/2019 KJE #169 catch protected cell issue
  // PAs can't edit TCID or date added column, so we need to ignore these when highlighting.
  // check only for access to all the other columns.
  */
  var statusCell = sheet.getRange(lineNumber+1,statusCol,1,1);
  var rest = sheet.getRange(lineNumber+1,dateAddedCol+1,1,sheet.getMaxColumns()-(dateAddedCol+1));
  var html = "";
  if(!statusCell.canEdit() || !rest.canEdit()){
    html = newText(html, robotName + " attempted to clear this employee on your behalf, but encountered an error. Please clear them by hand.",2);
    html = newText(html, "Sorry about that.",2);
    message.forward(myEmail.toLowerCase(),{htmlBody: html + "<br/><br/>- - -<br/><br/>" + message.getBody()});
    GmailApp.sendEmail("kjegerdal@madison.k12.wi.us", "Cleared employee error", "html", {
      htmlBody: "Row: " + sheet.getRange(lineNumber+1,1,1,sheet.getMaxColumns()).getDisplayValues() + "<br/><br/>" + message.getBody()});
    return; 
  }
  var email = row[emailCol-1];
  var firstName = row[firstNameCol-1];
  var lastName = row[lastNameCol-1];
  var prog = row[programCol-1];
  var pos = row[jobCodeCol-1];
  var sup = row[supCol-1];
  var pr = row[payrollApvsCol-1]; //*12/19/2019 KJE #242 enable %PAYROLL% token
  var startDate = row[startDateCol-1];
  var cc = row[CCCol-1];
  var curBNum = row[bNumCol-1];
  if((/\d{6}/).test(String(curBNum))){
    bNum = curBNum;
  }
  var bNumCell = range.getCell(lineNumber+1,bNumCol);
  //*6/6/2019 KJE clear AFTER setting bnum -- if error, won't result in cleared staff without bnum
  bNumCell.setValue(bNum);
  bNumCell.setFontWeight("normal");
  range.getCell(lineNumber+1,statusCol).setValue(cleared);
  setStatusColor(cleared,sheet,row,statusCell); //*10/8/2019 KJE #169 ignore TCID and date added cols
  setStatusColor(cleared,sheet,row,rest); //*10/8/2019 KJE #169 ignore TCID and date added cols
  var clearedEmailError = false;
  try{
    makeClearedEmail(firstName,pos,prog,bNum,sup,cc,email,pr,startDate); //*12/19/2019 KJE #242 enable %PAYROLL% token  *3/20/2024 KJE #513 enable start date token
  }catch(e){
    clearedEmailError = true;
  }
  html = newText(html, "This message was processed by " + robotName + ". " + firstName + " " + lastName + " was cleared in the SED and their B number was set to " + bNum + ".",2);
  if(!clearedEmailError){
    html = newText(html, "Additionally, in your drafts folder you'll find an e-mail notifying them they are cleared. It's up to you to send the e-mail.",2);
  }else{
    html = newText(html, "Additionally, " + robotName + " tried to draft an e-mail notifying them, but encountered an error. You can draft the email again by choosing their name in the SED, then using <i>Menu-PAs-Draft Cleared Email.</i>");
  }
  tryTryAgain(function(){ //*3/6/23 KJE hotfix
    message.forward(myEmail.toLowerCase(),{htmlBody: html + "<br/><br/>- - -<br/><br/>" + message.getBody()});
  });
}
function draftCleared(rows){
  try{
    getUi();
    makePleaseWait();
    getVals();
    if(sheet.getName() != hiredSheetName && sheet.getName() != activeSheetName){
      closePleaseWait();
      ui.alert("This option can only be used on the \"Hired\" or \"Active\" sheets.");
      return;
    }
    var foundOne = false;
    if(rows){
      for(var i = 0; i < rows.length; i++){
        var row = vals[rows[i]-1];
        if(row[statusCol-1] != cleared){
          continue; 
        }
        makeClearedEmail(row[firstNameCol-1],row[jobCodeCol-1],row[programCol-1],row[bNumCol-1],row[supCol-1],row[CCCol-1],row[emailCol-1],row[payrollApvsCol-1],row[startDateCol-1]); //*12/19/2019 KJE #242 enable %PAYROLL% token  *3/20/2024 KJE #513 enable start date token
        foundOne = true;
      }
    }else{
      var myEmail = getUser().getEmail();
      for(var i = hiredSheetLastHeader; i < vals.length; i++){
        var row = vals[i];
        if(row[statusCol-1] != cleared){
          continue; 
        }
        if(row[PACol-1].indexOf(myEmail) == -1){
          continue; 
        }
        makeClearedEmail(row[firstNameCol-1],row[jobCodeCol-1],row[programCol-1],row[bNumCol-1],row[supCol-1],row[CCCol-1],row[emailCol-1],row[payrollApvsCol-1],row[startDateCol-1]); //*12/19/2019 KJE #242 enable %PAYROLL% token  *3/20/2024 KJE #513 enable start date token
        foundOne = true;
      }
    }
    closePleaseWait();
    if(foundOne){
      ui.alert("Drafted at least one e-mail. Hooray!");
    }else{
      if(rows){
        ui.alert("Couldn't find any active staff in the lines you selected.");
      }else{
        ui.alert("Couldn't find any active staff under your name."); 
      }          
    }
  }catch(e){
    if((e.message).indexOf("Timed out") != -1){ //handle ui.alert timeout
      //do nothing
    }else if((e.message).indexOf("Invalid") != -1){
      ui.alert("Invalid email for either the staff, their supervisor, or the CC column. Please correct it and try again.");
    }else{
      GmailApp.sendEmail(SARLib.getErrorEmails(env), "draftCleared() Error", e + "\n" + e["stack"]);
      Logger.log(e.message + "\n" + e["stack"]);
      closePleaseWait();
      getUi().alert("Error processing your entry. An administrator has been notified.");
    }
  }
}
function draftClearedSelected(){
  draftCleared(getActiveRows());
}
function statusEmail(message,firstName,lastName,folder,id,processed,myEmail){
  var html = "";
  if(processed.length == 0){
    html = newText(html, robotName + " found " + firstName + " " + lastName + ", ID" + id + " in the SED, but none of the attached files were recognized or the employee didn't have a folder to put them in.",2);
  }else{
    html = newText(html, "This message was processed by " + robotName + ". The following files were found for ");
    html = newText(html, firstName + " " + lastName + ", ID" + id + ":",1);
    for(var i = 0; i < processed.length; i++){
      html = newText(html, "- " + processed[i],1); 
    }
    html = newLine(html);
    html = newText(html,"If the filename included a #, SEDbot marked the file as \"Done.\" If not, and it is, in fact, done, you’ll have to mark it as \"Done\" manually.",2);
  }
  html = newText(html,"If SEDbot didn’t recognize a file, you’ll have to rename it something more conspicuous, then run a file scan for SEDbot to recognize it." +
                 "\"Conspicuous\" names are in the notes of each column header. For instance, mouse over the column header for \"W4\" to see conspicuous names for a W4.",2);
  html = newText(html,"Here is a <a href=\"" + folder.getUrl() + "\">link</a> to the employee's folder.");
  try{
    tryTryAgain(function(){ //*5/6/20 KJE #253 replace manual try again with tryTryAgain
      message.forward(myEmail.toLowerCase(),{htmlBody: html + "<br/><br/>- - -<br/><br/>" + message.getBody()});
    });
  }catch(e){ //*5/6/20 KJE #253 catch too many attachments error
    if((e.message).indexOf("Limit Exceeded: Email Total Attachments Size") != -1){
      tryTryAgain(function(){ //*5/6/20 KJE #253 catch too many attachments error: send without attachments
        message.forward(myEmail.toLowerCase(),{
          htmlBody: html + "<br/><br/>- - -<br/><br/>" + message.getBody(),
          attachments: []
        });
      });
    }else{
      throw e; 
    }
  }
}
function getLastRun(fxnName){
  var prop = lastRunProp + fxnName;
  var lastRun = getUserProp(prop);
  if(!lastRun){
    return null;
  }
  return lastRun.substring(0,10); //Gmail search doesn't use MS, cut off last 3 digits
}
//*7/21/20 KJE #275 try again, handling data storage error
function setLastRun(fxnName,tryAgainAfter){
  var prop = lastRunProp + fxnName;
  //*7/21/20 KJE #275 try again, handling data storage error
  var list = getServerErrorList();
  list.push("Data storage error");
  tryTryAgain(function(){
    setUserProp(prop, String(new Date().valueOf())); //keep as a number or will auto-convert to scientific format :(
  },(tryAgainAfter ? tryAgainAfter : 5000),5,list);
}
function createReminders() {
  var todayMillis = (new Date()).getTime();
  getSSByID();
  getVals();
  getFormulas();
  var myEmail = getUser().getEmail();
  var madeOne = false;
  for(var i = hiredSheetLastHeader; i < vals.length; i++){
    var row = vals[i];
    if(row[statusCol-1] == cleared){continue;} //ignore cleared employees
    if(row[PACol-1].indexOf(myEmail) == -1){continue;} //only this person's employees
    if(!row[emailCol-1]){continue;} //only employees with e-mail addresses listed
    var lastReminder = row[reminderSentCol-1];
    if(!lastReminder){lastReminder = row[emailSentCol-1];}
    if(!lastReminder){continue;}
    lastReminder = new Date(lastReminder).valueOf();
    if(row[juliaCol-1]){continue;} //already sent for processing, ignore
    try{
      if(row[startDateCol-1]){
        if(row[startDateCol-1].valueOf() < todayMillis){ //if past start date
          if((todayMillis-lastReminder)/86400000 > 4){ //use every four days
            if(makeOneReminder(row,i,3,todayMillis)){
              madeOne = true;
            }
            continue;
          }
        }
      }
      if(!row[reminderSentCol-1]){ //never been sent a reminder -- use a week. 
        if((todayMillis-lastReminder)/86400000 > 7){ //every week
          if(makeOneReminder(row,i,1,todayMillis)){
            madeOne = true;
          }
          continue;
        }
      }else{ //previously sent a reminder
        if((todayMillis-lastReminder)/86400000 > 14){ //every two weeks
          if(makeOneReminder(row,i,2,todayMillis)){
            madeOne = true;
          }
          continue;
        }
      }
    }catch(e){
      if((e.message).indexOf("permission") != -1){
        return; //as long as this is automatic, we can just ignore errors and go on to the next one
        //if ever manually initiated, display errors to end users
      }
    }
  }
  if(madeOne){
    GmailApp.sendEmail(myEmail, robotName + " report", robotName + " made some reminders for your hires that haven't completed their paperwork yet. They are in your drafts folder.")
  }
}
/**
 * assumes range, sheet
 * 
 * 12/16/2021 MDH #330 pass env to getNeededDocsText
 */
function makeOneReminder(row,rowNum,type,todayMillis){
  var firstName = row[firstNameCol-1];
  var attachments = [];
  
  //special handling for returning agt only
  var recEmp = row[recEmpCol-1];
  recEmp = (recEmp && !isNaN(recEmp)) ? row[recEmpCol-1].getTime() : 0;
  if((todayMillis - recEmp) < (31540000000*0.25)){ //last employment less than three months ago
    //attach agt
    if(!row[signedSACol-1]){
      var EA;
      var EAform = formulas[rowNum][SACol-1];
      if(EAform){
        EA = DriveApp.getFileById(getIDFromHyperlink(EAform));
      }else{
        var today = Utilities.formatDate(new Date(), "CST", "MM-dd-YY");
        EA = makeOneAgreement(row,today,rowNum+1);
      }
      attachments.push(EA);
    }else{ //agt already returned
      return 0; 
    }
    //use firstName
    var replaceArray = {};
    replaceArray["FIRSTNAME"] = firstName;
    if(type == 3){ //*10/22/2019 KJE #168 don't say two weeks ago if past start date
      replaceArray["INTRO"] = "You are passed your start date for this position.";
    }else{
      replaceArray["INTRO"] = "About two weeks ago, I sent you a reminder about completing paperwork for your position at MSCR. Please remember that failure to complete and return your paperwork can affect your start date with us and result in a delay of pay.";
    }
    makeHtmlEmail(row[emailCol-1],row[CCCol-1],attachments,cannedRetrRem,replaceArray);
  }else{
    //not returning agt only -- manually build html
    var html = "";
    
    //Salutation
    html = newText(html,"Hi "+ firstName.trim() +",",2);
    
    //Intro
    if(type == 1){
      html = newText(html,"About a week ago, I e-mailed you about completing paperwork for your new job at MSCR.",2);
      // html = newText(html,"Please complete, digitally sign, and reply to this email with the following as attachments:",1); //*5/9/22 KJE #397 //4/29/23 YMC #437 not needed
    }else if(type == 2){
      html = newText(html,"About two weeks ago, I sent you a reminder about completing paperwork for your new job at MSCR.",2);
      // html = newText(html,"Please complete, digitally sign, and reply to this email with the following as attachments:",1); //*5/9/22 KJE #397 //4/29/23 YMC #437 not needed
    }else if(type == 3){
      html = newText(html,"Your start date with MSCR has passed, but you haven't completed all your paperwork. <b>We cannot pay you until you complete your paperwork & appointments.</b>",2); //4/29/23 YMC #437 update language
      // html = newText(html,"Please complete, digitally sign, and reply to this email with the following as attachments:",1); //*5/9/22 KJE #397 //4/29/23 YMC #437 not needed
    }
    
    var EAform = formulas[rowNum][SACol-1];
    if(EAform){
      var EAID = getIDFromHyperlink(EAform);
    }
    
    html = newText(html,SARLib.getNeededDocsText(row,EAID,attachments,(type == 3),env)); //*12/16/2021 MDH #330 pass env
    
    html = newText(html,"Best,",2);
    html = newText(html,getSignature()); //GetCannedResponses.getSignature()
    
    var subj;
    if(type != 3){
      subj = "MSCR: Paperwork needed before you can start! (ID" + row[idCol-1] + ")";
    }else{
      subj = "REMINDER: We cannot pay you until you complete your paperwork! -- MSCR (ID" + row[idCol-1] + ")"
    }
    try{
      GmailApp.createDraft(row[emailCol-1], subj, "Please view this e-mail on a device that renders html or reply to this e-mail requesting a plain text version.", {
        htmlBody: html,
        cc: row[CCCol-1],
        attachments: attachments
      });
    }catch(e){
      if((e.message).indexOf("Invalid") != -1){ //handle invalid address. Reminders are always behind-the-scenes, no need to alert.
        //do nothing
      }else if((e.message).indexOf("Service") != -1){
        Utilities.sleep(2000);
        GmailApp.createDraft(row[emailCol-1], subj, "Please view this e-mail on a device that renders html or reply to this e-mail requesting a plain text version.", {
          htmlBody: html,
          cc: row[CCCol-1],
          attachments: attachments
        });
      }else{
        throw e; 
      }
    }
  }
  
  //update last reminder sent column
  range.getCell(rowNum+1,reminderSentCol).setValue(new Date());
  
  //made an agreement
  return 1;
}
function toPAManual(){
  try{
    toPA(true);
  }catch(e){
    if((e.message).indexOf("No item with the given ID could be found, or you do not") != -1){
      getUi().alert("At least one document cannot be attached to an e-mail because you don't have access to it. Try checking each document for the employee you're trying to complete to make sure you have access. You will need to re-upload any documents you don't have access to, then run Menu > PAs > Check for Files to re-link them to the employee, then Menu > PAs > Check for Completion one more time to attach the new file." +
                    "\n\nSorry about that.");
    }else{
      throw e; 
    }
  }
}
function toPA(showUI,rows){
  try{
    if(showUI){
      getUi();
      makePleaseWait();
    }
    getVals();
    if(showUI && sheet.getName() != hiredSheetName){
      closePleaseWait();
      ui.alert("This option can only be used on the \"Hired\" sheet.");
      return;
    }
    getFormulas();
    var foundOne = false;
    var myEmail = getUser().getEmail();
    var PA = getPA(myEmail);
    if(rows){
      if(rows.length == 1){
        newCheckRowForFilesNoFolder(rows[0]-1);
        vals = null, formulas = null;
        getVals();
        getFormulas();
      }
      for(var i = 0; i < rows.length; i++){
        var rowNum = rows[i]-1;
        var row = getRowWithMRED(vals[rowNum],rowNum); //*6/15/23 KJE #465 force MRED to be filled before sending completion email
        try{ //*7/20/20 KJE #272 notify PAs when attachments are too large
          if(oneRowToPA(row,formulas[rowNum],PA,myEmail)){ //*5/13/2020 KJE #161 CC PA if different than sender //*6/15/23 KJE #465 force MRED to be filled before sending completion email
            //must be done row-by-row to avoid desynchronization when multiple PAs run this at night
            sheet.getRange(rowNum+1,juliaCol).setValue(new Date());
            sheet.getRange(rowNum+1,lastEditCol).setValue(new Date());
            foundOne = true;
          }
        }catch(e){ //*7/20/20 KJE #272 notify PAs when attachments are too large
          if((e.message).indexOf("Total Attachments Size") != -1){
            if(!showUI){
              //do nothing (continue) 
            }else{
              ui.alert("The attachments for employee " + vals[rowNum][SARLib.idCol] + " were too big to attach. If some of them are images, try saving them as pdfs, then try again."); 
            }
          }
        }
      }
    }else{
      for(var i = hiredSheetLastHeader; i < vals.length; i++){
        var row = vals[i];
        if(row[PACol-1].indexOf(myEmail) == -1){continue;} //only this person's employees
        row = getRowWithMRED(row,i); //*6/15/23 KJE #465 force MRED to be filled before sending completion email
        //*3/21/24 KJE #452 check for updated SA before attaching anything to an email
        if(showUI){ //!showUI is covered by newCheckForFiles in nightlyChecks
          newCheckForFiles(false,row);
        }
        try{ //*5/6/2020 KJE ignore access errors in nightlyChecks
          if(oneRowToPA(row,formulas[i],PA,myEmail)){ //*5/13/2020 KJE #161 CC PA if different than sender
            //must be done row-by-row to avoid desynchronization when multiple PAs run this at night
            sheet.getRange(i+1,juliaCol).setValue(new Date());
            sheet.getRange(i+1,lastEditCol).setValue(new Date());
            foundOne = true;
          }
        }catch(e){ //*5/6/2020 KJE ignore access errors in nightlyChecks //*7/20/20 KJE #272 ignore when attachments too large
          if(!showUI &&
             (e.message).indexOf("No item with the given ID could be found, or you do not have permission to access it.") != -1
          || (e.message).indexOf("Total Attachments Size") != -1){
            continue; 
          }else if(showUI){
            throw e; 
          }
        }
      }
    }
    if(showUI){
      closePleaseWait();
      if(foundOne){
        ui.alert("Drafted at least one e-mail. Hooray!");
      }else{
        if(rows){
          ui.alert("Couldn't find any completed staff in the lines you selected.\n\nThis ignores staff already sent to be cleared. To redo those staff, clear the \"Sent for Verif.\" column.\n\nIf you're sure a staff has all \"x\"s in their paperwork columns, run \"Check for Files,\" then try again.");
        }else{
          ui.alert("Couldn't find any completed staff under your name.\n\nThis ignores staff already sent to be cleared. To redo those staff, clear the \"Sent for Verif.\" column."); 
        }          
      }
    }else{
      if(foundOne){
        tryTryAgain(function(){ //*3/6/23 KJE hotfix
          GmailApp.sendEmail(myEmail, robotName + " report", robotName + " drafted some completion e-mails for you. They are in your drafts folder.");
        });
      }
    }
  }catch(e){
    if((e.message).indexOf("Timed out") != -1){
      return; 
    }
    GmailApp.sendEmail(SARLib.getErrorEmails(env), "toPA() Error", e + "\n" + e["stack"]);
    Logger.log(e.message + "\n" + e["stack"]);
    if(showUI){
      closePleaseWait();
      getUi().alert("Error processing your entry. An administrator has been notified.");
    }
  }
}
/**
 * Checks to see if a given row has a most recent employment date (including "not found"). If so, returns the row as-is. If not, gets it, fills out N/A columns as appropriate, then returns the N/A'd row.
 *
 * Initial version of this function is not very well optimized as it it assumed it will not be needed often. When not needed, it quits before doing any work.
 *
 * 6/15/2023 KJE #465 created
 *
 * @param {object} row The row in question, a list containing a list.
 * @param {Number} rowNum The row number of the row in question.
 * @author Kyle J. Egerdal
 */
function getRowWithMRED(row,rowNum){
  if(row[recEmpCol-1]){
    return row;
  }
  var MREDId = getMREDId(row);
  var needMRED = {};
  needMRED[MREDId] = "";
  var activeVals = getActiveSheet().getDataRange().getValues();
  for(var i = activeSheetLastHeader; i < activeVals.length; i++){
    checkMRED(activeVals[i],needMRED);
  }
  var archiveVals = getArchiveSheet().getDataRange().getDisplayValues();
  for(var i = SARLib.archiveLastHeader; i < archiveVals.length; i++){
    checkMRED(archiveVals[i],needMRED);
  }
  var newMRED = needMRED[MREDId];
  if(newMRED){ //exists and isn't empty
    var range = sheet.getRange(rowNum+1,recEmpCol);
    range.setValue(new Date(newMRED));
    range.setNumberFormat("M/d/yyyy");
    updateAgeAtStart(sheet,rowNum+1);
  }else{
    var tempRange = sheet.getRange(rowNum+1,recEmpCol)
    tempRange.clearDataValidations(); //*4/24/23 KJE #490 handle putting text in a date-formatted cell, which throws an error on next batch upload
    tempRange.setValue(MREDNotFound);
  }
  SpreadsheetApp.flush();
  return sheet.getRange(rowNum+1,1,1,sheet.getMaxColumns()).getValues()[0];
}
function toPASelected(){
  try{
    toPA(true,getActiveRows());
  }catch(e){
    if((e.message).indexOf("No item with the given ID could be found, or you do not") != -1){
      getUi().alert("At least one document cannot be attached to an e-mail because you don't have access to it. Try checking each document for the employee you're trying to complete to make sure you have access. You will need to re-upload any documents you don't have access to, then run Menu > PAs > Check for Files to re-link them to the employee, then Menu > PAs > Check for Completion one more time to attach the new file." +
                    "\n\nSorry about that.");
    }else{
      throw e; 
    }
  }
}
var t18Default = "Since they have worked for us before, please use the W4, WT4, direct deposit, and I9 already on file.";
/**
 * 5/13/2020 KJE #161 CC PA if different than sender
 * 2/25/22 MDH #382 Covid Vaccination
 * 3/1/22 MDH #383 for returning emps: only require if MRED before Nov. 1
 * 3/20/24 KJE #404 remove COVID column
 */
function oneRowToPA(row,rowFormulas,PA,myEmail){
  if(row[statusCol-1] == cleared){return false;} //ignore cleared employees
  if(row[juliaCol-1]){return false;} //already sent for processing, ignore
  var attachments = [];
  //ATSE
  var ATSE = getIDfromUrl(getUrl(rowFormulas[SACol-1]));
  if(ATSE && row[SACol]){
    attachments.push(ATSE);
  }else{return false;}
  //*3/21/24 KJE #404 attach approval if it's there
  Logger.log(row[approvalCol-1]);
  Logger.log(!row[approvalCol-1]);
  if(!row[approvalCol-1]){ //If column is blank, PA manually removed N/A, so we expect something, but there's nothing there. Quit out.
    return false;
  }
  var approval = getIDfromUrl(getUrl(rowFormulas[approvalCol-1]));
  if(approval){
    attachments.push(approval);
  }
  //skip W4, WT4, DD, and I9 for employees that turned 18 since their last employment less than a year ago.
  var recEmp = row[SARLib.recEmpCol-1];
  var lt1y = false,t18Statement = null;
  var hasBNum = (recEmp && recEmp != MREDNotFound);
  if(hasBNum){
    if((new Date().valueOf() - new Date(addMillenium(recEmp))) < 31540000000){ //3.154e10 ms in a year
      lt1y = true; 
    }
  }
  if(!lt1y){
    //W4, WT4
    var W4 = getIDfromUrl(getUrl(rowFormulas[W4Col-1]));
    var WT4 = getIDfromUrl(getUrl(rowFormulas[WT4Col-1]));
    if(W4 && row[W4Col]){
      if(WT4 && row[WT4Col]){
        if(W4[0] == WT4[0]){//W4 and WT4 are the same
          attachments.push(W4);
        }else{//they are different
          attachments.push(W4);
          attachments.push(WT4);
        }
      }else{return false;}
    }else{return false;}
    //DD
    var DD = getIDfromUrl(getUrl(rowFormulas[DDCol-1]));
    if(DD && row[DDCol]){
      attachments.push(DD);
    }else{return false;}
    //I9
    var I9 = getIDfromUrl(getUrl(rowFormulas[I9Col-1]));
    if(I9 && row[I9Col]){ //must be verified
      attachments.push(I9);
    }else{return false;}
  }else{
    t18Statement = t18Attachments(row,attachments,rowFormulas);
  }
  //FP/BG
  if(row[BCCol-1] != "N/A"){ //not everyone needs BG check
    if(!row[SARLib.DCFCol-1]){ //*10/8/2019 KJE #166 DCF
      return false; 
    }
    var BC = getIDfromUrl(getUrl(rowFormulas[BCCol-1]));
    if(BC && row[BCCol]){
      attachments.push(BC);
    }else{return false;}
  }else if(row[BCCheckCol-1] == ""){ //*3/20/24 KJE #404 BC double check column
    return false;
  }
  //TB
  if(row[TBCol-1] != "N/A"){ //not everyone needs TB test
    if(row[TBCol]){ //must be verified
      var TB = getIDfromUrl(getUrl(rowFormulas[TBCol-1]));
      if(TB && row[TBCol]){
        attachments.push(TB); 
      }else{return false;}
    }else{return false;}
  }
  //*3/20/24 KJE #404 remove COVID column
  //Covid
  // if(!row[covidVaxCol-1]){  //* 2/25/22 MDH #382 Covid Vaccination
  //   if(!hasBNum || (new Date(addMillenium(recEmp)) < new Date("11/01/2021"))){  //* 3/1/22 MDH #383 for returning emps: only require if MRED before Nov. 1
  //     return false;
  //   }
  // }
  //WP
  if(row[WPCol-1] != "N/A"){ //not everyone needs WP
    var WP = getIDfromUrl(getUrl(rowFormulas[WPCol-1]));
    if(WP){
      attachments.push(WP);
    }else{return false;}
  }
  //draft email
  var firstName = row[firstNameCol-1];
  var lastName = row[lastNameCol-1];
  var age = row[ageCol-1];
  var u18 = false,t18 = false;
  if(age != "18+"){
    if(age < 18){
      u18 = true; 
    }else{
      t18 = true; 
    }
  }
  var html = makeDoneBody(firstName, lastName, t18, t18Statement, u18, hasBNum, lt1y, row[startDateCol-1]);
  for(var i = 0; i < attachments.length; i++){ //now that we know we're sending it
    attachments[i] = DriveApp.getFileById(attachments[i]); //replace file IDs with actual files
  }
  var thisEmpsPA = row[PACol-1]; //*5/13/2020 KJE #161 CC PA if different than sender
  var cc = (thisEmpsPA.indexOf(myEmail) == -1) ? thisEmpsPA : ""; //*5/13/2020 KJE #161 CC PA if different than sender

  tryTryAgain(function(){ 
     GmailApp.createDraft(PA, (lt1y ? "RETURNING " : "NEW" ) + " - " + lastName + ", " + firstName + " (ID" + row[idCol-1] + "): Completed Paperwork",
                       "HTML is not enabled in your email client. Sad face!", {
                       htmlBody: html,
                       cc: cc, //*5/13/2020 KJE #161 CC PA if different than sender
                       attachments: attachments
                       }); 

   });
 
  return true;
}
function getPA(user){
  var listVals = tryTryAgain(function(){
    return getListSheet().getDataRange().getValues();
  });
  for(var i = listSheetLastHeader; i < listVals.length; i++){
    if(listVals[i][listSheetPACol-1].indexOf(user) != -1){
      return listVals[i][listSheetPAProcessorCol-1];
    }
  }
  return "";
}
function t18Attachments(row,attachments,rowFormulas){
  //W4, WT4
  var W4 = getIDfromUrl(getUrl(rowFormulas[W4Col-1]));
  var WT4 = getIDfromUrl(getUrl(rowFormulas[WT4Col-1]));
  var newSubmissions = ""
  if(W4){ //not dummy
    if(WT4){
      if(W4[0] == WT4[0]){//W4 and WT4 are the same
        attachments.push(W4);
      }else{//they are different
        attachments.push(W4);
        attachments.push(WT4);
      }
      newSubmissions = "W4 and WT4"
    }else{
      attachments.push(W4); 
      newSubmissions = "W4"
    }
  }else if(WT4){
    attachments.push(WT4);
    newSubmissions = "WT4"
  }
  //DD
  var DD = getIDfromUrl(getUrl(rowFormulas[DDCol-1]));
  if(DD){
    attachments.push(DD);
    if(newSubmissions != ""){
      newSubmissions = newSubmissions + " and "
    }
    newSubmissions = newSubmissions + "Direct Deposit form"
  }
  if(newSubmissions != ""){
    return "are choosing to submit a new " + newSubmissions + ", attached.";
  }else{
    return "";
  }
}

function makeDoneBody(firstName, lastName, t18, t18Statement, u18, hasBNum, lt1y, startDate){
  var html = newText("","Hello,",2); //Reminders.newText(html,text,newLines)
  html = newText(html,getNextAnimal(),2);
  html = newText(html,"Please find attached the completed paperwork for " + firstName + " " + lastName + ".",2);
  if(!lt1y){
    if((new Date()).valueOf() - startDate.valueOf() > daySecs*3){
      html = newText(html,"This employee's start date is at least three days ago. I <b>have/have not <font color=\"red\">(choose one)</font></b> verified their start date with their supervisor.",2);
    }
  }
  if(u18){
    html = newText(html,"This employee is under 18 and does not need a background check or TB test. ");
    if(lt1y && t18Statement){
      html = newText(html,"They have worked for us before and ");
    }
  }else if(t18 && lt1y){
    html = newText(html,"This employee has turned 18 since their last employment less than a year ago. ");
    if(t18Statement){
      html = newText(html,"They "); 
    }
  } 
  if(lt1y){
    html = newText(html," " + (t18Statement ? t18Statement : t18Default),2);
  }
  html = newText(html,"Let me know if anything is unsatisfactory or missing." + (hasBNum ? "" : " Otherwise, please reply with his/her B# when you have it."),2);
  html = newText(html,"Thanks!",1);
  html = newText(html,betterGetSignature()); //GetCannedResponses.getSignature()
  return html;
}
function getNextAnimal(){
  var cur = getProp("Animal");
  if(!cur){cur = 0;}
  cur = (Number(cur) + 1) % 6;
  setProp("Animal", cur);
  switch(cur){
    case 0:
      return getAnimalCode("1f408"); //cat
    case 1:
      return getAnimalCode("1f40f"); //sheep
    case 2:
      return getAnimalCode("1f410"); //goat
    case 3:
      return getAnimalCode("1f428"); //koala
    case 4:
      return getAnimalCode("1f422"); //turtle
    case 5:
      return getAnimalCode("1f416"); //pig
    default:
      return getAnimalCode("1f408"); //cat
  }
}
function getAnimalCode(id){
  return "<img src=\"https://mail.google.com/mail/e/" + id + "\">";
}
function makeClearedEmail(fname,pos,prog,bnum,sup,cc,email,payroll,startDate){
  var replaceArray = {};
  replaceArray["FIRSTNAME"] = fname;
  replaceArray["JOBCODE"] = pos;
  replaceArray["PROGRAM"] = prog;
  replaceArray["BNUMBER"] = bnum;
  replaceArray["SUPERVISOR"] = sup.substring(0,sup.indexOf("<")-1); //*12/19/2019 KJE #202
  replaceArray["PAYROLL"] = payroll.substring(0,sup.indexOf("<")-1); //*12/19/2019 KJE #242 enable %PAYROLL% token
  replaceArray["STARTDATE"] = Utilities.formatDate(startDate, "CST", "MM-dd-YY"); //*3/20/2024 KJE #513 enable start date token
  makeHtmlEmail(email,sup + "," + cc,null,clearedToWork,replaceArray); 
}