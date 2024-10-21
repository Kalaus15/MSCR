//Copyright (c) Kyle Egerdal 2018. All Rights reserved.

var env = "DEV";
var testing = true;
var daySecs = 86400000;

//----------------------------- COLUMNS AND ROWS -----------------------------//

var activeSheetLastHeader = SARLib.activeSheetLastHeader; //last header row
var ageCol = SARLib.ageCol;
var approvalCol = SARLib.approvalCol;
var birthdayCol = SARLib.birthdayCol;
var bNumCol = SARLib.bNumCol;
var BCCheckCol = SARLib.BCCheckCol; //*3/20/24 KJE #404 add BCCheckCol
var BCCol = SARLib.BCCol;
var CCCol = SARLib.CCCol;
var dateAddedCol = SARLib.dateAddedCol;
var DDCol = SARLib.DDCol;
var emailCol = SARLib.emailCol;
var emailSentCol = SARLib.emailSentCol;
var endDateCol = SARLib.endDateCol;
var firstDateLoggedCol = SARLib.firstDateLoggedCol;
var firstNameCol = SARLib.firstNameCol;
var folderCol = SARLib.folderCol;
var dateAddedCol = SARLib.dateAddedCol;
var directorCol = SARLib.directorCol;
var jobCodeCol = SARLib.jobCodeCol;
var hiredSheetLastHeader = SARLib.hiredSheetLastHeader;
var idCol = SARLib.idCol;
var I9Col = SARLib.I9Col;
var juliaCol = SARLib.juliaCol;
var lastEditCol = SARLib.lastEditCol;
var lastKeepCol = SARLib.lastKeepCol;
var lastNameCol = SARLib.lastNameCol;
var listSheetImmutableColumns = SARLib.listSheetImmutableCols;
var listSheetLastHeader = SARLib.listSheetLastHeader;
var listSheetMasterCol = SARLib.listSheetMasterCol;
var listSheetPACol = SARLib.listSheetPACol;
var listSheetPAProcessorCol = SARLib.listSheetPAProcessorCol;
var listSheetPARetCol = SARLib.listSheetPARetCol;
var listSheetPANewCol = SARLib.listSheetPANewCol;
var listSheetProcessorCol = SARLib.listSheetProcessorCol;
var listSheetProgCol = SARLib.listSheetProgCol;
var listSheetProgORCol = SARLib.listSheetProgORCol;
var listSheetSDCol = SARLib.listSheetSDCol;
var listSheetSiteCol = SARLib.listSheetSiteCol;
var listSheetSiteORCol = SARLib.listSheetSiteORCol;
var munisCodeCol = SARLib.munisCodeCol;
var munisSheetCodeCol = SARLib.munisSheetCodeCol;
var munisSheetJobCol = SARLib.munisSheetJobCol;
var munisSheetLastHeader = SARLib.munisSheetLastHeader;
var munisSheetName = SARLib.munisSheetName;
var munisSheetProgramCol = SARLib.munisSheetProgramCol;
var munisSheetSiteCol = SARLib.munisSheetSiteCol;
var notesCol = SARLib.notesCol;
var PACol = SARLib.PACol;
var payRateCol = SARLib.payRateCol;
var payrollApvsCol = SARLib.payrollApvsCol;
var phCol = SARLib.phCol;
var photoCol = SARLib.photoCol;
var PRSheetCodeCol = SARLib.PRSheetCodeCol;
var PRSheetPRStartCol = SARLib.PRSheetPRStartCol;
var programCol = SARLib.programCol;
var quitDateCol = SARLib.quitDateCol;
var recEmpCol = SARLib.recEmpCol;
var reminderSentCol = SARLib.reminderSentCol;
var SACol = SARLib.SACol;
var signedSACol = SARLib.signedSACol;
var siteCol = SARLib.siteCol;
var startDateCol = SARLib.startDateCol;
var statusCol = SARLib.statusCol;
var supCol = SARLib.supCol;
var t18Col = SARLib.t18Col;
var TBCol = SARLib.TBCol;
var txtOkCol = SARLib.txtOkCol;
var W4Col = SARLib.W4Col;
var WPCol = SARLib.WPCol;
var WT4Col = SARLib.WT4Col;
//var covidVaxCol = SARLib.covidVaxCol; //* 2/25/22 MDH #382 Covid Vaccination *3/20/24 KJE #404 remove COVID Column
//
var allowPasteMin = lastNameCol;
var allowPasteMax = txtOkCol;

//-------------------------------- CONSTANTS -------------------------------//

var tz = SARLib.tz;
var cleared = SARLib.cleared;
var quit = SARLib.quit;
var missingInfo = SARLib.missingInfo;
var deleteMe = SARLib.deleteMe;
var replaced = SARLib.replaced; //*11/6/2021 MDH #336 add replaced
var hiredStatusList = [cleared,quit,missingInfo,deleteMe,replaced] //*11/6/2021 MDH #336 add replaced

//-------------------------------- FOLDERS --------------------------------//

var empFolder;
var empFolderID = "143Rdjd8S7BO0e7hkETtCymnCk_je5Zt_";
var empFolderPRDID = "1TZwl2UJhDSf-0zR9XhLOj07D_phvUAvo";
var archiveFolder;
var archiveFolderID = "1PwD6Cbjeq08tADXPjaNZrSkKcBCsC_s_";
var archiveFolderPRDID = "1eaCDAIiXwxUhkpl5yzRtUf1WShf1D0LO";

function getEmpFolder(){
  if(!empFolder){
    if(env == "PRD"){
      empFolder = DriveApp.getFolderById(empFolderPRDID);
    }else{
      empFolder = DriveApp.getFolderById(empFolderID);
    }
  }
  return empFolder;
}

function getArchiveFolder(){
  if(!archiveFolder){
    if(env == "PRD"){
      empFolder = DriveApp.getFolderById(archiveFolderPRDID);
    }else{
      archiveFolder = DriveApp.getFolderById(archiveFolderID);
    }
  }
  return archiveFolder;
}

function getCannedResponsesFolder(){
  return DriveApp.getFolderById("1YnOoHteXoSl0J624zr9Hbp7XkeYA6jaP");
}

function getPACSVsFolder(){
  return DriveApp.getFolderById("1hw-rzIEk1tFNQabZ6hV09ZpNlP29RSoZ");
}

//-------------------------------- UI ------------------------------------//

var ui;

function getUi(){
  if(!ui){
    ui = SpreadsheetApp.getUi();
  }
  return ui;
}

//-------------------------------- USERS ---------------------------------//

var user;
var email;

function getUser(){
  if(!user){
    user = Session.getActiveUser();
  }
  return user;
}

//-------------------------------- PATTERNS -------------------------------//

var jCodeREx = /M\d{3}/;

//------------------------------ PROPERTIES -------------------------------//

var sp;
var up;

function getSP(){
  if(!sp){
    sp = PropertiesService.getScriptProperties();
  }
  return sp;
}

function getProp(prop){
  return getSP().getProperty(prop);
}

function setProp(prop,val){
  getSP().setProperty(prop,val);
}
function getUserProperties(){
  if(!up){
    up = PropertiesService.getUserProperties();
  }
  return up;
}

function getUserProp(prop){
  return tryTryAgain(function(){
    return getUserProperties().getProperty(prop);
  });
}

function setUserProp(prop,val){
  tryTryAgain(function(){
    getUserProperties().setProperty(prop,val);
  });
}

//-------------------------------- SHEETS ---------------------------------//

var cols;
var activeSheet;
var activeSheetName = SARLib.activeSheetName;
var activeSheetFormulas;
var archiveSheet;
var archiveSheetName = "Archive";
var formulas;
var hiredSheet;
var hiredSheetName = SARLib.hiredSheetName;
var listSheetName = SARLib.listSheetName;
var listSheet;
var munisSheet;
var munisSheetName = "MUNIS";
var payRateSheet;
var payRateSheetName = "Pay Scale";
var range;
var rows;
var sheet;
var ss;
var vals;
var listSheetVals; //*3/11/23 KJE #445

function getActiveSheet(){
  if(!activeSheet){
    activeSheet = getSS().getSheetByName(activeSheetName);
  }
  return activeSheet;
}

function getArchiveSheet(){
  if(!archiveSheet){
    archiveSheet = getSS().getSheetByName(archiveSheetName);
  }
  return archiveSheet;
}

function getDimensions(){
  getVals();
  if(!rows){
    rows = vals.length;
  }
  if(!cols){
    cols = vals[0].length;
  }
}

/**
 * 1/10/2022 MDH #350 allow forceNew
 */
function getFormulas(forceNew){  //*1/10/2022 MDH #350 allow forceNew
  if(!formulas || forceNew){
    formulas = getRng().getFormulas();
  }
  return formulas;
}

function getHiredSheet(){
  if(!hiredSheet){
    hiredSheet = getSS().getSheetByName(hiredSheetName);
  }
  return hiredSheet;
}

function getListSheet(){
  if(!listSheet){
    listSheet = getSS().getSheetByName(listSheetName);
  }
  return listSheet;
}
//*3/11/23 KJE #445
function getListSheetVals(){
  if(!listSheetVals){
    getSSByID();
    listSheetVals = getListSheet().getDataRange().getDisplayValues();
  }
  return listSheetVals;
}

function getMUNISSheet(){
  if(!munisSheet){
    munisSheet = getSS().getSheetByName(munisSheetName);
  }
  return munisSheet;
}

function getPayRateSheet(){
  if(!payRateSheet){
    payRateSheet = getSS().getSheetByName(payRateSheetName);
  }
  return payRateSheet;
}

function getRng(){
  if(!range){
    range = getSht().getDataRange();
  }
  return range;
}

function getSht(){
  if(!sheet){
    sheet = tryTryAgain(function(){ //*11/3 KJE tryTryAgain
      return getSS().getActiveSheet();
    });
  }
  return sheet;
}

function getSS(){
  if(!ss){
    ss = tryTryAgain(function(){
      return SpreadsheetApp.getActiveSpreadsheet();
    }); //*3/4/20 KJE #225 fix unexpected error
  }
  return ss;
}
var TSTIDProp = "TSTID";
function getSSByID(){
  return tryTryAgain(getSSByIDWrapped,null,10); //*3/11/23 KJE hotfix 3 -> 10 tries
}
function getSSByIDWrapped(){
  if(!ss){
    switch(env){
      case "DEV":
        ss = SpreadsheetApp.openById("12AIZ1eMfbPNdk4SOFp9CvJR4JoPQ7BGs2MpMXwaGtn8");
        break;
      case "TST":
        ss = SpreadsheetApp.openById(getProp(TSTIDProp));
        break;
      case "PRD":
        ss = SpreadsheetApp.openById("1Fi3xZQ-pHDQ-U6RbRIjK24eJrOJjpYQQWnQQpHZ7Qyk");
        break;
    }
  }
  return ss;
}

function getVals(){
  if(!vals){
    vals = tryTryAgain(function(){ //*6/14/2021 KJE tryTryAgain
      return getRng().getValues();
    });
  }
  return vals;
}

function getDispVals(){
  if(!vals){
    vals = tryTryAgain(function(){ //*6/14/2021 KJE tryTryAgain
      return getRng().getDisplayValues();
    });
  }
  return vals;
}
