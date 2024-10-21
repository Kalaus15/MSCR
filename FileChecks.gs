//Copyright (c) Kyle Egerdal 2018. All Rights reserved.

function newCheckForFiles(){
  newCheckForFiles(false);
}
function newCheckForFilesManual(){
  newCheckForFiles(true); 
}
function newCheckForFilesSelected(){
  newCheckForFiles(true,getActiveRows()); 
}
function newCheckForFiles(showUI,rows){
  var email = getUser().getEmail();
  var key = email + "/newCheckForFiles";
  if(!SARLib.PAReadLock(key,env)){return;}
  var error;
  try{
    if(showUI){
      makePleaseWait();
    }
    email = email.toUpperCase();
    getVals(); //Globals.gs
    if(showUI && sheet.getName() != hiredSheetName){
      SARLib.PAReadUnlock(key,env);
      closePleaseWait();
      getUi().alert("This option can only be used on the \"Hired\" sheet.");
      return;
    }
    getFormulas();
    var foundOne;
    if(rows){
      for(var i = 0; i < rows.length; i++){
        var row = rows[i]-1;
        if(newCheckRowForFilesNoFolder(row)){
          foundOne = true;
          //must be done row-by-row to avoid desynchronization when multiple PAs run this at night
          range.getCell(row+1,lastEditCol).setValue(new Date());
        }
      }
    }else{
      for(var row = hiredSheetLastHeader; row < vals.length; row++){
        if(vals[row][PACol-1].toUpperCase().indexOf(email) == -1){continue;} //only this PA's employees
        if(newCheckRowForFilesNoFolder(row)){
          foundOne = true;
          //must be done row-by-row to avoid desynchronization when multiple PAs run this at night
          range.getCell(row+1,lastEditCol).setValue(new Date());
        }
      }
    }
  }catch(e){
    error = e;
  }
  SARLib.PAReadUnlock(key,env);
  if(showUI){
    closePleaseWait();
  }
  if(error){
    GmailApp.sendEmail(SARLib.getErrorEmails(env), "newCheckForFiles() Error", error + error["stack"]);
    throw error; 
  }
  if(showUI){
    if(foundOne){
      SpreadsheetApp.flush();
      Browser.msgBox("Found at least one new file. Hooray!");
    }else{
      var htmlOutput = HtmlService
      .createHtmlOutput("No files found. Need help? Click <a href=\"https://docs.google.com/document/d/1sYuXMHjp_S8Ppc1JJBcRlTBoH00QtBy2UsNuwHDkppo/edit#heading=h.pmyk98y1kw7v\" target=\"_blank\">here</a>.")
      .setHeight(30)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, "File Search");
    }
  }
}
//assumes: values
function newCheckRowForFilesNoFolder(row){
  var folder = getFolder(row); //HelperFxns.gs
  return newCheckRowForFiles(row,folder);
}
function newCheckRowForFiles(row,folder){
  var files=folder.getFiles();
  var firstName = vals[row][firstNameCol-1];
  var lastName = vals[row][lastNameCol-1];
  var foundOne = false;
  var orderedFiles = [],found = [];
  while (files.hasNext()) {
    var file = files.next();
    var date = file.getLastUpdated();
    var put = false;
    for(var i = 0; i < orderedFiles.length; i++){
      if(orderedFiles[i].getLastUpdated() < date){
        orderedFiles.splice(i,0,file);
        put = true;
        break;
      }
    }
    if(!put){
      orderedFiles.push(file); 
    }
  }
  for(var i = 0; i < orderedFiles.length; i++){
    if(newMatchFile(firstName,lastName,row,orderedFiles[i],found) != -1){
      foundOne = true;
    }
  }
  return foundOne;
}

//agreement (always assume it's signed)
var SANameEx = /SIGNED|ATSE|AGREEMENT|\bAGT|\bSA\b|CONTRACT/; //*10/22/2019 KJE #183 remove -
var SAName = "SA";
var SAPre = 1;
//*3/21/24 KJE #519 approval
var apvEx = /APPROV/;
var apvName = "Approval";
//tax forms (both w4 and wt4)
var tfEx = /TAX FORMS|\bTF|TAXFORMS/; //starts with TF or contains space TF. Prevents matching on names like "Atford"
var tfPre = 2;
var tfName = "TF";
//w4
var w4Ex = /W4|W-4/;
var w4Pre = 2;
var w4Name = "W4";
//wt4
var wt4Ex = /WT4|WT-4|WT_4/;
var wt4Pre = 2;
var wt4Name = "WT4";
//dd
var ddEx = /\bDD|DIRECT/; //starts with dd or contains space dd. Prevents matching on names with two d's like "Kaddo"
var ddPre = 3;
var ddName = "DD";
//i-9 and typos
var i9Ex = /I9|I-9|I_9|1-9/;
var i9Pre = 4;
var i9Name = "I9";
//background check
var bgEx = /\bBG|\bBC|\bFP/; //prevents matching tricky names that probably exist
var bgPre = "";
var bgName = "BC";
//tb test
var tbEx = /\bTB/;
var tbPre = "";
var tbName = "TB";
//work permit
var wpEx = /\bWP|WORK PERMIT/;
var wpPre = 5;
var wpName = "WP";
//*3/20/24 KJE #404 remove COVID column
//covid vaccine  //* 3/1/22 MDH #383 SEDbot detect covid vaccine 
// var covidEx = /COVID|VACCINE|VAX/;
// var covidPre = ""
// var covidName = "COVID";

/**
 * 3/1/22 MDH #383 SEDbot detect covid vaccine
 * 3/20/24 KJE #404 Remove COVID Column
 * 3/21/24 KJE #519 supervisor approval
 */
function newMatchFile(firstName,lastName,row,file,found){
  var rowFormulas = formulas[row];
  file = myFile(file); //make sure we own it
  var name = file.getName().toUpperCase();
  var markDone = (name.indexOf("#") != -1) ? true : false;
  var formula="=HYPERLINK(\""+file.getUrl()+"\",\"x\")";
  if(SANameEx.test(name)){//agreement
    if(found.indexOf(SAName) == -1){
      if(rowFormulas[SACol-1] != formula){
        //must be done row-by-row to avoid desynchronization when multiple PAs run this at night
        range.getCell(row+1,SACol).setFormula(formula);
        if(markDone){range.getCell(row+1,SACol+1).setValue("x");}
        setName(SAPre,firstName,lastName,SAName,file);
        found.push(SAName + (markDone ? " (done)" : ""));
        return SAName;
      }else{
        found.push(SAName + (markDone ? " (done)" : ""));
        return -1;
      }
    }else{
      return -1;
    }
  }
  //*3/21/24 KJE #519 approval
  if(checkOneName(apvEx,name,row,approvalCol-1,formula,null,firstName,lastName,apvName,file,found,markDone)){return apvName;}
  if(tfEx.test(name)){//both tax forms in one file
    if(found.indexOf(tfName) == -1){
      if(rowFormulas[W4Col-1] != formula || rowFormulas[WT4Col-1] != formula){
        //must be done row-by-row to avoid desynchronization when multiple PAs run this at night
        range.getCell(row+1,W4Col).setFormula(formula);
        if(markDone){range.getCell(row+1,W4Col+1).setValue("x");}
        range.getCell(row+1,WT4Col).setFormula(formula);
        if(markDone){range.getCell(row+1,WT4Col+1).setValue("x");}
        setName(tfPre,firstName,lastName,tfName,file);
        found.push(tfName + (markDone ? " (done)" : ""));
        return tfName;
      }else{
        found.push(tfName + (markDone ? " (done)" : ""));
        return -1;
      }
    }else{
      return -1; 
    }
  }else{
    if(w4Ex.test(name)){//w4
      if(wt4Ex.test(name)){ //W4 and WT4 are same file
        if(rowFormulas[W4Col-1] != formula && rowFormulas[WT4Col-1] != formula){
          if(found.indexOf(tfName) == -1){
            setName(tfPre,firstName,lastName,tfName,file);
            //must be done row-by-row to avoid desynchronization when multiple PAs run this at night
            range.getCell(row+1,WT4Col).setFormula(formula);
            if(markDone){range.getCell(row+1,WT4Col+1).setValue("x");}
            found.push(tfName + (markDone ? " (done)" : ""));
            return tfName;
          }else{
            return -1; 
          }
        }else{
          found.push(tfName + (markDone ? " (done)" : ""));
          return -1;
        }
      }else{
        if(found.indexOf(w4Name) == -1 && found.indexOf(tfName) == -1){
          if(rowFormulas[W4Col-1] != formula){
            setName(w4Pre,firstName,lastName,w4Name,file);
            //must be done row-by-row to avoid desynchronization when multiple PAs run this at night
            range.getCell(row+1,W4Col).setFormula(formula);
            if(markDone){range.getCell(row+1,WT4Col+1).setValue("x");}
            found.push(w4Name + (markDone ? " (done)" : ""));
            return w4Name;
          }else{
            found.push(w4Name + (markDone ? " (done)" : ""));
            return -1;
          }
        }else{
          return -1; 
        }
      }
    }
    //wt4 only
    if(found.indexOf(tfName) == -1){
      if(checkOneName(wt4Ex,name,row,WT4Col-1,formula,wt4Pre,firstName,lastName,wt4Name,file,found,markDone)){return wt4Name;}
    }
  }
  //everything else
  if(checkOneName(ddEx,name,row,DDCol-1,formula,ddPre,firstName,lastName,ddName,file,found,markDone)){return ddName;}
  if(checkOneName(i9Ex,name,row,I9Col-1,formula,i9Pre,firstName,lastName,i9Name,file,found,markDone)){return i9Name;}
  //*10/2/2019 KJE #166 DCF
  if(checkOneName(bgEx,name,row,BCCol-1,formula,bgPre,firstName,lastName,bgName,file,found,markDone)){
    var DCFCell = range.getCell(row+1,SARLib.DCFCol);
    if(DCFCell.getValue() != "N/A"){
      DCFCell.setValue("x");
    }
    return bgName;
  }
  if(checkOneName(tbEx,name,row,TBCol-1,formula,tbPre,firstName,lastName,tbName,file,found,markDone)){return tbName;}
  //*3/20/24 KJE #404 Remove COVID Column
  //if(checkOneName(covidEx,name,row,covidVaxCol-1,formula,covidPre,firstName,lastName,covidName,file,found,markDone)){return covidName;} //* 3/1/22 MDH #383 SEDbot detect covid vaccine 
  if(checkOneName(wpEx,name,row,WPCol-1,formula,wpPre,firstName,lastName,wpName,file,found,false)){return wpName;} //*6/18/2019 KJE WPs are never done
  return -1;
}
function checkOneName(ex,nowName,row,col,formula,pre,first,last,thenName,file,found,markDone){
  if(found.indexOf(thenName) == -1){
    if(ex.test(nowName)){
      if(formulas[row][col] != formula){
        //must be done row-by-row to avoid desynchronization when multiple PAs run this at night
        range.getCell(row+1,col+1).setFormula(formula);
        if(markDone){range.getCell(row+1,col+2).setValue("x");}
        setName(pre,first,last,thenName,file);
        found.push(thenName + (markDone ? " (done)" : ""));
        return true;
      }else{
        found.push(thenName + (markDone ? " (done)" : ""));
      }
    }
  }
  return false;
}

var dummyID = "1qDUIXYT4leBftnaZioCRIaXfk55-05fK";

function newCheckNames(){
  var startRow=getDummyCell()[0]+1;
  var endRow=getLastCell()[0];
  for(var row=startRow;row < endRow; row++){
    var firstName = formulas[row][firstNameCol];
    var lastName = formulas[row][lastNameCol];
    //W4, WT4
    var W4 = getIDfromCell(row,W4Col);
    var WT4 = getIDfromCell(row,WT4Col);
    if(W4 && W4[0] != dummyID){ //not dummy
      if(WT4 && WT4[0] != dummyID){ //not dummy
        if(W4 == WT4){ //W4 and WT4 are same file
          setName(firstName,lastName,tfName,DriveApp.getFileById(W4));
        }else{
          setName(firstName,lastName,w4Name,DriveApp.getFileById(W4));
          setName(firstName,lastName,wt4Name,DriveApp.getFileById(WT4));
        }
      }else{
        setName(firstName,lastName,w4Name,DriveApp.getFileById(W4)); 
      }
    }else if(WT4 && WT4[0] != dummyID){ //not dummy
      setName(firstName,lastName,wt4Name,DriveApp.getFileById(WT4));
    }
    setOneName(row,SACol-1,firstName,lastName,SAName);
    setOneName(row,signedSACol-1,firstName,lastName,signedSAName);
    setOneName(row,WPCol-1,firstName,lastName,wpName);
    setOneName(row,DDCol-1,firstName,lastName,ddName);
    setOneName(row,I9Col-1,firstName,lastName,i9Name);
    setOneName(row,BGCol-1,firstName,lastName,bgName);
    setOneName(row,TBCol-1,firstName,lastName,tbName);
    setOneName(row,WPCol-1,firstName,lastName,wpName);
  }
}
function setOneName(row,col,first,last,name){
  var id = getIDfromCell(row,col);
  if(id){
    if(id[0] != dummyID){
      setName(first,last,name,DriveApp.getFileById(id));
    }
  }
}
function myFile(file){
  getUser();
  if(file.getOwner() != user){ //not ours, shared by a user on Google Drive
    var it = file.getParents();
    while(it.hasNext()){
      var parent = it.next();
      var owner = tryTryAgain(function(){ //*11/14/2020 KJE tryTryAgain
        return parent.getOwner(); 
      });
      if(owner == user){
        var myFile = file.makeCopy(file.getName(), parent);
        file.setTrashed(true);
        return myFile;
      }
    }
  }
  return file;
}

var extRegex = /.[^.]+$/; //a period, at least one of anything but a period, the end of the string

function setName(pre,firstName,lastName,type,file){
  var ext = extRegex.exec(file.getName())[0]; //get extension
  if(!ext){ext = "";}
  pre = pre ? pre + " " : ""; //*3/21/24 KJE #519 passing "null" actually put the word "null"
  return file.setName(pre + lastName + ", " + firstName + " - " + type + ext);
}
function massUploadInstr(){
  SpreadsheetApp.getUi().alert("This utility will attempt to file everything in the mass upload folder and send e-mails to the appropriate PAs notifying them it's done so. Those e-mails will come from you. Then, it will email you to let you know it's done so." +
                               "\n\nAnything successfully filed will be removed from the folder. Anything left in the folder was not filed -- the e-mail will tell you why." +
                               "\n\nFiles should be named like so: <LNAME, FNAME - FILETYPE>.\n\nFiletypes can be abbreviated using the lexicon in the notes of each file type's column. For instance, if you mouse over the column header for \"Agreements,\" you will see the allowable titles for that file type." +
                              "\n\nAdd a pound sign (#) to file names to auto-check the \"Done\" columns.");
}
function scheduleMassUpload(){
  if(cancelMassUpload(true,true)){
    getUi().alert("A mass upload is already scheduled or running. Only one can be run at a time.");
    return; 
  }
  var masterEmail = getUser().getEmail();
  setProp("massUploadEmail",masterEmail);
  ScriptApp.newTrigger("massUpload").timeBased().after(60000).create();
  getUi().alert("A new mass upload will begin in 60 seconds. Results will be sent to " + masterEmail + ".\n\nPlease allow 30 minutes for this process to complete.");
}
function cancelMassUpload(noUi,checkOnly){
  var triggers = ScriptApp.getProjectTriggers();
  for(var i = 0; i < triggers.length; i++){
    var trigger = triggers[i];
    if(trigger.getHandlerFunction() == "massUpload"){
      if(checkOnly){
        return true; 
      }
      ScriptApp.deleteTrigger(triggers[i]);
      break;
    }
  }
  if(!noUi){
    SpreadsheetApp.getUi().alert("Any scheduled uploads have been cancelled.\n\nAny already running cannot be stopped. Please wait for them to finish before starting more.");
  }
  if(checkOnly){
    return false; 
  }
}
function massUpload(){
  try{
    var masterEmail = getProp("massUploadEmail");
    var statusEmailSubject = "Mass Upload " + Utilities.formatDate(new Date(), "CST", "MM/dd/yyyy hh:mm a");
    var timer = String(new Date().valueOf());
    var folderId = null;
    switch(env){
      case "DEV":
      case "TST":
        folderId = "1hbkYpFBJpksnjCqwrbWZCqXe3znSyixP";
        break;
      case "PRD":
        folderId = "1A-CTUY0M8GhaXtxrX-HwvZO89Pi0ziBL";
        break;
    }
    var folder = DriveApp.getFolderById(folderId);
    if(!folder){
      GmailApp.sendEmail(masterEmail, statusEmailSubject, "Couldn't find the mass upload folder. Contact your developer to have this fixed.");
      cancelMassUpload(true);
      return;
    }
    range = getHiredSheet().getDataRange();
    vals = range.getDisplayValues();
    formulas = range.getFormulas();
    var files = folder.getFiles();
    var names = {};
    var closeMatches = {};
    var pas = {};
    var errors = {};
    var timedOut = false;
    while(files.hasNext()){
      if(String(new Date().valueOf()) - timer > 60*25*1000){ //25' timeout -- leave time to send e-mails!
        timedOut = true;
        break; 
      }
      var file = files.next();
      massUploadOneFile(file,folder,names,closeMatches,pas,errors);
    }
    makeMassUploadEmail(masterEmail,statusEmailSubject,timedOut,pas,closeMatches,errors);
    cancelMassUpload(true);
  }catch(e){
    GmailApp.sendEmail(SARLib.getErrorEmails(env), "massUpload() error", e.message + "\n" + e["stack"]);
    cancelMassUpload(true);
    GmailApp.sendEmail(masterEmail,statusEmailSubject,"An error occured. Developer has been notified.\n\nSorry about that.");
  }
}
function makeMassUploadEmail(masterEmail,statusEmailSubject,timedOut,pas,closeMatches,errors){
  var html = "";
  html = SARLib.newText(html, "Hello,", 2);
  html = SARLib.newText(html, "You recently ran a mass upload from the SED.", 0);
  if(timedOut){
    html = SARLib.newText(html, " Unforunately, it timed out before all files could be processed. Try again to process the remaining files.", 2);
    html = SARLib.newText(html, "Files that were processed are below.", 2);
  }else{
    html = SARLib.newLines(html, 2);
  }
  html = SARLib.newText(html, "Successfully filed:", 1);
  var oneSuccess = false;
  for(pa in pas){
    oneSuccess = true;
    html = SARLib.newText(html, "- PA: " + pa, 1);
    var paText = "";
    paText = SARLib.newText(paText, "Hello,", 2);
    paText = SARLib.newText(paText, "The following files were recently filed in the SED for you by an adminstrator:", 1);
    for(name in pas[pa]){
      html = SARLib.newText(html, "-- " + name + ": " + pas[pa][name].reduce(function(a, b){return a.concat(", " + b)}), 1);
      paText = SARLib.newText(paText, "-- " + name + ": " + pas[pa][name].reduce(function(a, b){return a.concat(", " + b)}), 1);
    }
    paText = SARLib.newLines(paText, 1);
    paText = SARLib.newText(paText, "Thank you,", 2);
    paText = SARLib.newText(paText, "SEDbot,", 2);
    var sendToPa = (testing ? SARLib.getErrorEmails(env) : pa);
    try{
      GmailApp.sendEmail(sendToPa, "Documents filed for you in the SED", "Please view in a browser that renders html.", {
        htmlBody: paText
      });
    }catch(e){
      if((e.message).indexOf("server") != -1 || (e.message).indexOf("Service") != -1){
        Utilities.sleep(2000);
        GmailApp.sendEmail(sendToPa, "Documents filed for you in the SED", "Please view in a browser that renders html.", {
          htmlBody: paText
        });
      }else{
        throw e; 
      }
    }
  }
  if(!oneSuccess){
    html = SARLib.newText(html, "(none)", 1);
  }
  html = SARLib.newLines(html, 1);
  html = SARLib.newText(html, "Errors:", 1);
  var oneFailure = false;
  for(name in closeMatches){
    oneFailure = true;
    html = SARLib.newText(html, "- All files for " + name + ": couldn't find a staff with this name. Close matches: ");
    var list = closeMatches[name];
    html = SARLib.newText(html, (list.length > 0) ? list.reduce(function(a, b){return a.concat("; " + b)}) : "(none)", 1);
  }
  for(name in errors){
    oneFailure = true;
    html = SARLib.newText(html, "- File named \"" + name + "\": " + errors[name],1);
  }
  if(!oneFailure){
    html = SARLib.newText(html, "(none)", 1);
  }
  html = SARLib.newLines(html, 1);
  html = SARLib.newText(html, "Thank you,", 1);
  html = SARLib.newText(html, "SEDbot");
  GmailApp.sendEmail(masterEmail, statusEmailSubject, "Please view in a browser that renders HTML.",{
    htmlBody: html
  });
}
function massUploadOneFile(file,folder,names,allCloseMatches,pas,errors){
  var fileName = file.getName();
  try{
    var dot = fileName.indexOf(".");
    if(dot != -1){
      fileName = fileName.substring(0,dot); 
    }
    var parts = fileName.split(" - ");
    var fileType = parts[1].trim();
    var empName = parts[0].trim();
    var empNames = empName.split(",");
    var fName = empNames[1].trim();
    var fNameUC = fName.toUpperCase();
    var lName = empNames[0].trim();
    var lNameUC = lName.toUpperCase();
  }catch(e){
    errors[fileName] = "Make sure the file name matches the format [LNAME, FNAME - FILETYPE]. Spaces on either end of the - are imperative.";
    return;
  }
  if(allCloseMatches[empName]){
    return;
  }
  var closeMatches = [];
  if(names[empName]){
    moveOneFileForMassUpload(file,fileName,folder,empName,fName,lName,names[empName]["rowNum"],names[empName]["pa"],pas,errors);
    return;
  }else{
    for(var i = SARLib.hiredSheetLastHeader; i < vals.length; i++){
      var row = vals[i];
      if(row[SARLib.statusCol-1] == SARLib.deleteMe){continue;}
      var thisFName = row[SARLib.firstNameCol-1].trim().toUpperCase();
      var thisLName = row[SARLib.lastNameCol-1].trim().toUpperCase();
      if(!thisFName || !thisLName){continue;}
      if(!isMatch(fNameUC,lNameUC,thisFName,thisLName)){
        if(isCloseMatch(fNameUC,lNameUC,thisFName,thisLName)){
          closeMatches.push(cap(thisLName) + ", " + cap(thisFName)); 
        }
        continue;
      }
      var PA = row[SARLib.PACol-1];
      if(!PA){
        errors[fileName] = "Found ID#" + row[SARLib.idCol-1] + " in the SED, but couldn't find their PA.";
      }
      names[empName] = {};
      names[empName]["rowNum"] = i;
      names[empName]["pa"] = PA;
      moveOneFileForMassUpload(file,fileName,folder,empName,fName,lName,i,PA,pas,errors);
      return;
    }
  }
  allCloseMatches[empName] = closeMatches;
}
function moveOneFileForMassUpload(file,fileName,folder,empName,fName,lName,rowNum,pa,pas,errors){
  var found = [];
  if(newMatchFile(fName,lName,rowNum,file,found)){
    var empFolder = getFolder(rowNum);
    addFileCertainly(empFolder,file); //*7/17/19 #118 KJE folder doesn't always know it has file, even if file knows it has folder
    if(!hasFile(empFolder,file)){ //*7/17/19 #118 KJE folder doesn't always know it has file, even if file knows it has folder
      errors[fileName] = "Filed, but couldn't move the file to the employee's folder. Please do so by hand."; 
    }else{
      folder.removeFile(file);
    }
    if(pas[pa] === undefined){
      pas[pa] = {};
    }
    if(pas[pa][empName] === undefined){
      pas[pa][empName] = [];
    }
    pas[pa][empName].push(found[0]);
  }else{
    errors[fileName] = "Found this employee in row " + (rowNum + 1) + ", but couldn't identify the file type.";
  }
}
//*7/17/19 #118 KJE folder doesn't always know it has file, even if file knows it has folder
/**
 * *1/10/2022 MDH #350 deprecate addFile
 */
function addFileCertainly(folder,file){
  file.moveTo(folder); //*1/10/2022 MDH #350 deprecate addFile
  var tries = 0;
  while(!hasFile(folder,file) && tries < 10){
    Utilities.sleep(500);
    file.moveTo(folder); //*1/10/2022 MDH #350 deprecate addFile
    tries++;
  }
}
//*7/17/19 #118 KJE folder doesn't always know it has file, even if file knows it has folder
function hasFile(folder,file){
  var files = folder.getFiles();
  while(files.hasNext()){
    var curFile = files.next();
    if(curFile.getId() == file.getId()){
      return true; 
    }
  }
  return false;
}
function isMatch(fName,lName,fName2,lName2){
  return (fName == fName2 && lName == lName2);
}
var sharedLetterPercent = 0.7;
function isCloseMatch(fName,lName,fName2,lName2){
  //check word matches
  var words = fName.split(" ").concat(lName.split(" "));
  var words2 = fName2.split(" ").concat(lName2.split(" "));
  for(var i = 0; i < words.length; i++){
    if(words2.indexOf(words[i]) != -1){
      return true; 
    }
  }
  for(var i = 0; i < words2.length; i++){
    if(words.indexOf(words2[i]) != -1){
      return true; 
    }
  }
  //check character matches
  var similarCharacters = 0;
  for(var i = 0; i < fName.length; i++){
    if(fName2.indexOf(fName[i]) != -1){
      similarCharacters++; 
    }
  }
  for(var i = 0; i < lName.length; i++){
    if(lName2.indexOf(lName[i]) != -1){
      similarCharacters++; 
    }
  }
  if(similarCharacters/(fName.length + lName.length) >= sharedLetterPercent){
    return true; 
  }
  return false;
}