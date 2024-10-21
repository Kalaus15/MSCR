//Copyright (c) Kyle Egerdal 2018. All rights reserved.

/*
* ASSUMES:
* - MUNIS is alphabetical by Program Name
*/

var chooseSite = "Choose a site first.";
var chooseProgram = "Choose a program first.";
var chooseJob = "Choose a job first.";
var EIDProp = "EID";
var newEmpKey = "NEW";
var grandfatherPayRateColors = "#e6b8af,";//comma-delimited
var payRateIgnoreColors = "#666666,";//comma-delimited * 1/24/22 MDH #371 Ignore grey background cells
var munisIgnoreColors = "#999999," ;//comma-delimited
var sdBackgroundColor = "#ccfeff";
var pleaseWaitTitle = "Please Wait...";
var pleaseWaitHeight = 20;
var yearEx = /\/\d{2}$/;

/**
 * Handles permissions for the last row. See comments on the function setRowPermissions for why this is neccesary.
 * 
 * Not to be conflated with the simple trigger onChange. As an installable trigger this runs under the authorization of the user who owns the trigger, which means it can edit the sheet freely. If it were a simple trigger it would run under the authorization of the user editing the sheet, who doesn't have a lot of permissions (given the way we've set things up).
 * 
 * So, this should be set up as an installable trigger that runs when a user changes the spreadsheet.
 * 
 * As a general rule, this function should NOT throw any errors, but either (a) notify the end user they made a mistake, or (b) notify the end user an error occured, notify the developer of the error via email, and exit gracefully.
 * 
 * 11/9/2021 MDH #336 replaced functionality
 * 3/2/22 MDH #388 tryTryAgain
 * 
 * @param {object} event The Google Event object automatically passed when a user changes a spreadsheet.
 * @author Kyle Egerdal
 */
function myOnChange(event){
  try{
    tryTryAgain(function(){  //*3/2/22 MDH #388 tryTryAgain
      setRowPermissions();
    },null,10); //*7/7/22 KJE hotfix, 3 tries -> 10
  }catch(e){
    if((e.message).indexOf("Service error") != -1){
      Utilities.sleep(5000);
      myOnChange(event);
      return;
    }else if(e.message == "Sheet not found"){ //Google sometimes "can't find" sheets that are actually there. Get it next time.
      return;
    }
    GmailApp.sendEmail(SARLib.getErrorEmails(env), "myOnChange() Error", e.message + "\n" + e["stack"]);
  }
}
/**
 * The simple, automatic trigger (does not need to be set up in the "Triggers" menu) for when a user edits the spreadsheet. As a simple trigger, this script is run with the permissions of the user making the edit, which is severely limiting (as such, the majority of the automatic changes to the sheet are made with the installable trigger myOnEdit).
 * 
 * If the user is trying to clear an employee and that employee can't be cleared, does nothing. If they are cleared, auto-drafts an email to that person's paperwork processor regarding the employee being cleared (function PAEmails.makeClearedEmail).
 * 
 * @param {object} event The Google Event object automatically passed when a user edits a spreadsheet.
 * @author Kyle Egerdal
 */
function onEdit(event){
  try{
    ss = event.source;
    var sheet = ss.getActiveSheet();

    //---------------------------- HIRED SHEET ----------------------------//

    if(sheet.getName() == hiredSheetName){//"Hired"
      var range = event.range;
      if(range.getNumRows() != 1 || range.getNumColumns() != 1){ //editing more than one cell
        return;
      }
      var row = range.getRow();
      var col = range.getColumn();
      if(row <= hiredSheetLastHeader){ //ignore changes to header
        return;
      }if(col == statusCol){ //status was updated, validate if cleared and edit formatting for the entire row
        var status = event.value;
        if(status == cleared){
          var clearedError = empCanBeCleared(row,sheet);
          if(clearedError){
            return;
          }
          var rowData = sheet.getRange(row,1,1,sheet.getMaxColumns()).getDisplayValues()[0];
          makeClearedEmail(rowData[firstNameCol-1],rowData[jobCodeCol-1],rowData[programCol-1],rowData[bNumCol-1],rowData[supCol-1],rowData[CCCol-1],rowData[emailCol-1]);
        }
      }
    }
  }catch(e){
    GmailApp.sendEmail(SARLib.getErrorEmails(env), "onEdit() Error", e + "\n" + e["stack"]);
  }
}
/**
 * Main trigger for automatically making changes when users edit the SED. Handles, among other things:
 * - Editing permissions of columns to make sure supervisors and PAs can edit the columns they should be able to edit, and not others
 * - Formatting (ex. date cells)
 * - Data validation (ex. email addresses; B#s; drop-downs for program selection once site is selected, job selection once program is selected)
 * - Auto-filling MUNIS code, pay rate, site director, PA, etc. once enough information is known to do that 
 * - Validating whether staff are ready to be cleared
 * - Color coding (columns supervisors need to edit, rows based on status)
 * - Paperwork columns based on whether staff is returning or not, and their age
 * - Pop-ups notifying the editor of definite or likely mistakes
 * 
 * Not to be conflated with the simple trigger onEdit. As an installable trigger this runs under the authorization of the user who owns the trigger, which means it can edit the sheet freely. If it were a simple trigger it would run under the authorization of the user editing the sheet, who doesn't have a lot of permissions (given the way we've set things up).
 * 
 * So, this should be set up as an installable trigger that runs when a user edits the spreadsheet.
 * 
 * As a general rule, this function should NOT throw any errors, but either (a) notify the end user they made a mistake, or (b) notify the end user an error occured, notify the developer of the error via email, and exit gracefully.
 * 
 * @param {object} event The Google Event object automatically passed when a user edits a spreadsheet.
 * @author Kyle Egerdal
 */
function myOnEdit(event){
  try{
    var range = event.range;
    var row = tryTryAgain(function(){return range.getRow();});
    var col = tryTryAgain(function(){return range.getColumn();});
    ss = event.source;
    var sheet = tryTryAgain(function(){return ss.getActiveSheet();});

//---------------------------- HIRED SHEET ----------------------------//

    if(sheet.getName() == hiredSheetName){//"Hired"
      if(range.getNumRows() != 1 || range.getNumColumns() != 1){ //*3/27/19 KJE handle updateAgeAtStart if editing more than one cell
        var maxCol = col + range.getNumColumns()-1;
        if((col <= startDateCol && maxCol >= startDateCol) || (col <= birthdayCol && maxCol >= birthdayCol)){
          makePleaseWait();
          var maxRow = row + range.getNumRows()-1;
          for(var i = row; i <= maxRow; i++){
            updateAgeAtStart(sheet,i);
          }
          closePleaseWait();
        }
        return;
      } //end editing more than one cell
      if(row <= hiredSheetLastHeader){ //ignore changes to header
        return;
      }
      if(!col){col = range.getColumn();}
      if(!sheet.getRange(row, idCol).getValue()){ //no ID#, this employee is (presumably) new
        makePleaseWait();
        try{
          tryTryAgain(function(){ //*4/28/21 try again on service error
            fillNewRow(sheet,row);
          });
        }catch(e){
          if((e.message).indexOf("TCID") != -1){
            SpreadsheetApp.getUi().alert(e.message);
          }else{
            throw e;
          }
        }
        closePleaseWait();
      }
      if(col == idCol){
        var value = event.value;
        if(/^#\d+KJE$/.test(value)){
          range.setValue(value.substring(0,value.length-3));
          return;
        }
        var oldValue = event.oldValue;
        range.setValue(oldValue ? oldValue : "");
        getUi().alert("Only administrators can change TCIDs.");
        return;
      }else if(col == statusCol){ //status was updated, validate if cleared and edit formatting for the entire row
        var status = event.value;
        makePleaseWait();
        var rowRange = sheet.getRange(row,1,1,sheet.getMaxColumns());
        setBorders(sheet,row,rowRange);
        if(!status){
          sheet.getRange(row,quitDateCol).setValue(""); //clear quit date, if quit was on accident
        }
        if(status == cleared){
          var clearedError = empCanBeCleared(row,sheet);
          if(clearedError){
            var oldValue = event.oldValue;
            range.setValue(oldValue ? oldValue : "");
            closePleaseWait();
            getUi().alert("This employee cannot be cleared. They need:" + clearedError);
            return;
          }
        }else if((status == quit) || (status == replaced)){ //*11/9/2021 MDH #336 replaced functionality
          sheet.getRange(row,quitDateCol).setValue(new Date());
        }
        setStatusColor(status,sheet,row,rowRange);
        closePleaseWait();
      }else if(col == firstNameCol || col == lastNameCol){ //first or last name, ensure proper capitalization
        var name = event.value;
        if(name){
          makePleaseWait();
          range.setNumberFormat("@"); //*9/12/20 KJE #282 force to text (for last names like "True")
          range.setValue(cap(event.value)); //HelperFxns.gs
          closePleaseWait();
        }
      }else if(col == birthdayCol){ //change in birthday can change age at start
        var birthday = event.value;
        if(birthday){
          birthday = getDateCell(sheet,row,birthdayCol);
          var apxAge = (new Date() - new Date(birthday))/31540000000;
          makePleaseWait();
          try{
            updateAgeAtStart(sheet,row,birthday);
          }catch(e){ //*KJE 11/14 handle birthday being a variant of 18+
            if((e.message).indexOf("18") != -1){
              if(birthday.indexOf("18") != -1 && birthday.length < 4){
                birthday = "18+";
                sheet.getRange(row, col).setValue(birthday);
                updateAgeAtStart(sheet,row);
              }else{
                throw e;
              }
            }else{
              throw e;
            }
          }
          if(apxAge < 14){
            getUi().alert("The birthday you entered makes this staff " + (Math.round(apxAge*100)/100) + " years old. Please double-check it.");
          }
          closePleaseWait();
        }
      }else if(col == bNumCol){ //change in b# can impact PA (new vs returning). also set formatting
        if(event.value){
          makePleaseWait();
          var thisRange = sheet.getRange(row,col);
          if(event.value.indexOf(newEmpKey) != -1){
            thisRange.setFontWeight("bold");
          }else{
            thisRange.setFontWeight("normal");
          }
          thisRange.setHorizontalAlignment("right");
          var program = sheet.getRange(row,programCol).getValue();
          if(program){
            var listSheetVals = getListSheet().getDataRange().getValues();
            setPA(sheet,row,program,listSheetVals);
          }
          closePleaseWait();
        }
      }else if(col == emailCol){ //*4/16/2020 KJE #245 reduce errors from incorrect e-mail formatting
        var email = event.value;
        if(email){
          makePleaseWait();
          range.setValue(email.trim());
          closePleaseWait();
        }
      }else if(col == siteCol){ //site set - put in program
        makePleaseWait();
        var site = event.value;
        var progCell = sheet.getRange(row,programCol,1,1);
        var jcCell = sheet.getRange(row,jobCodeCol,1,1);
        if(!site){
          progCell.clearDataValidations();
          progCell.setValue(chooseSite);
          jcCell.clearDataValidations();
          jcCell.setValue(chooseProgram);
          clearedJobCodeCol(sheet,row);
          closePleaseWait();
          return;
        }else if(progCell.getValue() == chooseSite || site != event.oldValue){
          progCell.setValue("");
          jcCell.setValue("");
          clearedJobCodeCol(sheet,row);
        }
        setSup(sheet,row,site,null);
        var munisData = tryTryAgain(function(){return getMUNISSheet().getDataRange();});
        var munisVals = tryTryAgain(function(){return munisData.getValues();});
        var munisColors = tryTryAgain(function(){return munisData.getBackgrounds();});
        var progs = [];
        var alls = 0;
        var i = SARLib.munisSheetLastHeader; //*8/5/19 KJE deprecate munisSheetLastAll
        for(; i < munisVals.length; i++){ //*5/14/19 KJE don't hard-code number of alls
          if(munisVals[i][munisSheetSiteCol-1] != "!All"){
            break;
          }
          var val = munisVals[i][munisSheetProgramCol-1];
          if(progs.indexOf(val) == -1){
            progs.push(val);
            alls++;
          }
        }
        if(i >= munisVals.length){
          i = SARLib.munisSheetLastHeader;
        }
        var start = false;
        //MUNIS is a long list and it's alphabetical. Once we start matching, we can stop looking as soon as we stop matching.
        //That is, if the list is A A B B C C and we're looking for a B, once we get to the first B, we can stop as soon as we hit not B.
        for(; i < munisVals.length; i++){ //*8/5/19 KJE don't hard-code number of alls
          if(munisVals[i][munisSheetSiteCol-1] == site){ //cols are 1-indexed, values are 0-indexed
            start = true;
            if(munisIgnoreColors.indexOf(munisColors[i][munisSheetSiteCol-1]) != -1){continue};
            var val = munisVals[i][munisSheetProgramCol-1];
            if(progs.indexOf(val) == -1){
              progs.push(val);
            }
          }else{
            if(start){
              break;
            }
          }
        }
        var progRule = SpreadsheetApp.newDataValidation().requireValueInList(progs, true).setAllowInvalid(false).build();
        progCell.setDataValidation(progRule);
        if(progs.length == alls+1){
          var program = progs[alls];
          progCell.setValue(program);
          settedProgram(sheet,row,program,jcCell);
        }else{
          jcCell.clearDataValidations();
          jcCell.setValue(chooseProgram);
        }
        closePleaseWait();
      }else if(col == programCol){ //program was edited, we can now validate job code, supervisor, and PA
        makePleaseWait();
        var program = event.value;
        var jcCell = sheet.getRange(row,jobCodeCol,1,1);
        if(!program){
          jcCell.clearDataValidations();
          jcCell.setValue(chooseProgram);
          clearedJobCodeCol(sheet,row);
          closePleaseWait();
          return;
        }else if(jcCell.getValue() == chooseProgram || program != event.oldValue){
          jcCell.setValue("");
          clearedJobCodeCol(sheet,row);
        }
        settedProgram(sheet,row,program,jcCell);
        closePleaseWait();
      }else if(col == jobCodeCol){ //job code was edited, we can now get MUNIS code and pay rates
        makePleaseWait();
        var job = event.value;
        if(!job || job == chooseProgram){
          clearedJobCodeCol(sheet,row);
          closePleaseWait();
          return;
        }else if(job != event.oldValue){
          clearedJobCodeCol(sheet,row);
        }
        settedJob(sheet,row,job);
        closePleaseWait();
      }else if(col == startDateCol){
        var startDate = event.value;
        if(startDate){
          makePleaseWait();
          updateAgeAtStart(sheet,row);
          closePleaseWait();
        }
      }else if(col == endDateCol){
        var endDate = event.value;
        if(endDate){
          makePleaseWait();
          endDate = sheet.getRange(row,col).getDisplayValue();
          var asDate = new Date(endDate);
          if(asDate.valueOf() < (new Date().valueOf())){
            var newYear = asDate.getFullYear();
            if(newYear < 2000){
              newYear = newYear + 100;
            }else{
              newYear++;
            }
            asDate.setFullYear(newYear);
            var suggestion = Utilities.formatDate(asDate, tz, "MM/dd/YYYY");
            getUi();
            var resp = ui.alert("End Date",endDate + " is in the past. Did you mean " + suggestion + "?" +
                                "\n\n\"Yes\" - Replace " + endDate + " with " + suggestion +
                                "\n\"No\" - Keep " + endDate +
                                "\n\"Cancel\" - Delete " + endDate + " and enter something else", ui.ButtonSet.YES_NO_CANCEL);
            if(resp == ui.Button.YES){
              sheet.getRange(row,col).setValue(suggestion);
            //no -- keep as is
            }else if(resp == ui.Button.CANCEL){
              sheet.getRange(row,col).setValue("");
            }
          }
          closePleaseWait();
        }
        //*5/13/2020 KJE #241 consolidate PA notifications
//      }else if(col == PACol){ //PA was manually changed
//        makePleaseWait();
//        var PA = event.value;
//        if(PA){
//          notifyPADialogue(PA,sheet.getRange(row,idCol).getValue());
//        }
//        closePleaseWait();
      }else if(col == ageCol){ //*3/7/19 KJE updateWillBe18 if Age At Start input manually
        makePleaseWait();
        updateWillBe18(sheet,row);
        closePleaseWait();
      }else if(col == recEmpCol){ //most recent employment date
        makePleaseWait();
        updateWillBe18(sheet,row);
        closePleaseWait();
      }


//---------------------------- ACTIVE SHEET ----------------------------//


    }else if(sheet.getName() == activeSheetName){
      if(!row){row = range.getRow();}
      if(row <= hiredSheetLastHeader){ //ignore changes to header
        return;
      }
      if(!col){col = range.getColumn();}
      if(col == statusCol){
        makePleaseWait();
        var status = event.value;
        var rowRange = sheet.getRange(row,1,1,sheet.getMaxColumns());
        if(col == statusCol){
          if(!status || status == cleared){
            sheet.getRange(row,quitDateCol).setValue(""); //clear quit date, if quit was on accident
          }
          if((status == quit) || (status == replaced)){ //*11/9/2021 MDH #336 replaced functionality
            sheet.getRange(row,quitDateCol).setValue(new Date());
          }
        }
        setStatusColor(status,sheet,row,rowRange);
        closePleaseWait();
      }else if(col == siteCol){
        makePleaseWait();
        var changeTo = event.oldValue ? event.oldValue : ""; //will actually set to "undefined" if oldValue is undefined
        //if(!event.value){ //event.value always returning undefined for some reason. uncomment this once Google figures itself out
        //  sheet.getRange(row,col).setValue(changeTo);
        //}else{
          getUi();
          var resp = ui.alert("Site Change for Active Employee",
                              "Does this employee also need a new job code?",ui.ButtonSet.YES_NO_CANCEL);
          if(resp == ui.Button.YES){
            sheet.getRange(row,col).setValue(changeTo);
            ui.alert("For new job codes, please de-activate the current entry, then complete and clear a new one on the \"" + hiredSheetName + "\" tab.");
            //no -- keep as is
          }else if(resp == ui.Button.CANCEL){
            sheet.getRange(row,col).setValue(changeTo);
          }
        //}
        closePleaseWait();
      }

//---------------------------- LIST SHEET ----------------------------//


    }else if(sheet.getName() == listSheetName){ //Sites, SDs, PAs
      if(range.getNumRows() != 1 || range.getNumColumns() != 1){
        return;
      }
      var row = range.getRow();
      if(row == listSheetLastHeader){ //ignore changes to header
        return;
      }
      var col = range.getColumn();
      if(col == listSheetMasterCol){
        makePleaseWait();
        updateMasters(sheet.getDataRange().getValues());
        closePleaseWait();
      }else if(col == listSheetSDCol){
        makePleaseWait();
        updateSDs(sheet.getDataRange().getValues());
        sheet.getRange(2,col,sheet.getMaxRows()-1,2).sort(col);
        closePleaseWait();
      }else if(col == SARLib.listSheetSDSupsCol){
        makePleaseWait();
        updateAdmins(sheet.getDataRange().getValues());
        closePleaseWait();
      }else if(col == listSheetPACol){
        makePleaseWait();
        updatePAs(sheet.getDataRange().getValues());
        sheet.getRange(2,col,sheet.getMaxRows()-1,2).sort(col);
        closePleaseWait();
      }else if(col == listSheetSiteORCol){
        listSheetReplacement(payrollApvsCol,event.value,siteCol,sheet.getRange(row,listSheetSiteCol).getValue(),null,null,true);
      }else if(col == listSheetProgORCol){
        listSheetReplacement(supCol,event.value,programCol,sheet.getRange(row,listSheetProgCol).getValue());
      }else if(col == listSheetPARetCol){
        listSheetReplacement(PACol,event.value,programCol,sheet.getRange(row,listSheetProgCol).getValue(),bNumCol,/\d{6}/);
      }else if(col == listSheetPANewCol){
        listSheetReplacement(PACol,event.value,programCol,sheet.getRange(row,listSheetProgCol).getValue(),bNumCol,/NEW/);
      }


//---------------------------- MUNIS SHEET ----------------------------//


    }else if(sheet.getName() == munisSheetName){
      if(!row){row = range.getRow();}
      if(row <= munisSheetLastHeader){ //ignore changes to header
        return;
      }
      if(!col){col = range.getColumn();}
      if(col == munisSheetSiteCol){
        var site = event.value;
        if(!site){return;}
        var old = event.oldValue;
        if(!old){return;}
        if(site == old){return;}
        makePleaseWait();
        getListSheet();
        var listRange = listSheet.getRange(listSheetLastHeader+1,listSheetSiteCol,listSheet.getLastRow()-listSheetLastHeader,1);
        var listVals = listRange.getValues();
        for(var i = 0; i < listVals.length; i++){
          if(listVals[i][0] == old){
            listVals[i][0] = site;
            listRange.setValues(listVals);
            break;
          }
        }
        getHiredSheet();
        var hiredRange = hiredSheet.getRange(hiredSheetLastHeader+1,siteCol,hiredSheet.getMaxRows()-hiredSheetLastHeader,1);
        var hiredVals = hiredRange.getValues();
        var replacedOne = false;
        for(var i = 0; i < hiredVals.length; i++){
          if(hiredVals[i][0] == old){
            hiredVals[i][0] = site;
            replacedOne = true;
          }
        }
        if(replacedOne){
          hiredRange.setValues(hiredVals);
        }
        closePleaseWait();
      }else if(col == munisSheetProgramCol){
        var prog = event.value;
        if(!prog){return;}
        var old = event.oldValue;
        if(!old){return;}
        if(prog == old){return;}
        makePleaseWait();
        getListSheet();
        var listRange = listSheet.getRange(listSheetLastHeader+1,listSheetProgCol,listSheet.getLastRow()-listSheetLastHeader,1);
        var listVals = listRange.getValues();
        for(var i = 0; i < listVals.length; i++){
          if(listVals[i][0] == old){
            listVals[i][0] = prog;
            listRange.setValues(listVals);
            break;
          }
        }
        getHiredSheet();
        var hiredRange = hiredSheet.getRange(hiredSheetLastHeader+1,programCol,hiredSheet.getMaxRows()-hiredSheetLastHeader,1);
        var hiredValidations = hiredRange.getDataValidations();
        var hiredVals = hiredRange.getValues();
        var replacedOne = false;
        for(var i = 0; i < hiredVals.length; i++){
          if(hiredVals[i][0] == old){
            hiredVals[i][0] = prog;
            var validation = hiredValidations[i][0];
            if(validation){
              var criteria = validation.getCriteriaValues();
              criteria.splice(criteria.indexOf(old),1,prog);
              hiredValidations[i][0] = SpreadsheetApp.newDataValidation().requireValueInList(criteria, true).setAllowInvalid(false).build();
              replacedOne = true;
            }
          }
        }
        if(replacedOne){
          hiredRange.setDataValidations(hiredValidations);
          hiredRange.setValues(hiredVals);
        }
        closePleaseWait();
      }else if(col == munisSheetCodeCol){
        var code = event.value;
        if(!code){return;}
        var old = event.oldValue;
        if(!old){return;}
        if(code == old){return;}
        var site = sheet.getRange(row,munisSheetSiteCol).getValue();
        var prog = sheet.getRange(row,munisSheetProgramCol).getValue();
        var job = sheet.getRange(row,munisSheetJobCol).getValue();
        makePleaseWait();
        getHiredSheet();
        var hiredRange = hiredSheet.getRange(hiredSheetLastHeader+1,siteCol,hiredSheet.getMaxRows()-hiredSheetLastHeader,4);
        var hiredVals = hiredRange.getValues();
        var replacedOne = false;
        for(var i = 0; i < hiredVals.length; i++){
          if(hiredVals[i][0] == site){
            if(hiredVals[i][1] == prog){
              if(hiredVals[i][2] == job){
                if(hiredVals[i][3] == old){
                  hiredVals[i][3] = code;
                  replacedOne = true;
                }
              }
            }
          }
        }
        if(replacedOne){
          hiredRange.setValues(hiredVals);
        }
        closePleaseWait();
      }else if(col == munisSheetJobCol){
        var job = event.value;
        if(!job){return;}
        var old = event.oldValue;
        if(!old){return;}
        if(job == old){return;}
        var site = sheet.getRange(row,munisSheetSiteCol).getValue();
        var prog = sheet.getRange(row,munisSheetProgramCol).getValue();
        makePleaseWait();
        getHiredSheet();
        var hiredRange = hiredSheet.getRange(hiredSheetLastHeader+1,siteCol,hiredSheet.getMaxRows()-hiredSheetLastHeader,3);
        var hiredValidations = hiredRange.getDataValidations();
        var hiredVals = hiredRange.getValues();
        var replacedOne = false;
        for(var i = 0; i < hiredVals.length; i++){
          if(hiredVals[i][0] == site){
            if(hiredVals[i][1] == prog){
              if(hiredVals[i][2] == old){
                hiredVals[i][2] = job;
                var validation = hiredValidations[i][2];
                if(validation){
                  var criteria = validation.getCriteriaValues();
                  criteria.splice(criteria.indexOf(old),1,job);
                  hiredValidations[i][2] = SpreadsheetApp.newDataValidation().requireValueInList(criteria, true).setAllowInvalid(false).build();
                  replacedOne = true;
                }
              }
            }
          }
        }
        if(replacedOne){
          hiredRange.setDataValidations(hiredValidations);
          hiredRange.setValues(hiredVals);
        }
        closePleaseWait();
      }
    }
  }catch(e){
    if(e.message == "Timed out waiting for user response"){return;} //not helpful
    if((e.message).indexOf("Service Spreadsheets failed while accessing document with id 1Fi3xZQ-pHDQ-U6RbRIjK24eJrOJjpYQQWnQQpHZ7Qyk.") != -1){
      Utilities.sleep(1000);
      myOnEdit(event); //*4/9/20 KJE try again if can't access SED (this spreadsheet!)
      return; //*5/23/21 KJE don't display this error since we handled it
    }
    //*5/16/19 KJE handle .gs errors as (e).indexOf instead of (e.message).indexOf
    if((e.message).indexOf("Must be") != -1 || (e.message).indexOf("data validation") != -1){return;} //ideally we could tell the user to fix this but for now, ignore
    try{
      var sheet = event.source.getActiveSheet();
      var sheetName = sheet.getName();
      var range = event.range;
      var row = range.getRow();
      var col = range.getColumn();
      GmailApp.sendEmail(SARLib.getErrorEmails(env), "myOnEdit() Error (" + sheet.getParent().getName() + ")", e + "\n" + e["stack"] +
                         "\n\nSheet/Row/Col/Old/New\n" + sheetName + "/" + row + "/" + col + "/" + event.oldValue + "/" + event.value);
    }catch(e2){
      GmailApp.sendEmail(SARLib.getErrorEmails(env), "myOnEdit() Error", e + "\n" + e["stack"]);
      GmailApp.sendEmail(SARLib.getErrorEmails(env), "myOnEdit() Error", e2 + "\n" + e2["stack"]);
    }
    if(range.getNumRows() != 1 || range.getNumColumns() != 1){ //only handle editing of single cells
      return;
    }
    Logger.log(e.message + "\n" + e["stack"]);
    closePleaseWait();
    getUi().alert("Error processing your entry. An administrator has been notified.");
  }
}
function listSheetReplacement(thisCol,thisVal,ifCol,ifVal,andCol,regEx,allowBlank){
  if(!thisCol || (!thisVal && !allowBlank) || !ifCol || !ifVal){
    return;
  }
  if(!thisVal){
    thisVal = "";
  }
  makePleaseWait();
  pushListsToSheet(thisCol,thisVal,ifCol,ifVal,getHiredSheet(),hiredSheetLastHeader,andCol,regEx);
  pushListsToSheet(thisCol,thisVal,ifCol,ifVal,getActiveSheet(),activeSheetLastHeader,andCol,regEx);
  SpreadsheetApp.flush();
  closePleaseWait();
}
function pushListsToSheet(thisCol,thisVal,ifCol,ifVal,sheet,lastHeader,andCol,regEx){
  var vals = sheet.getDataRange().getDisplayValues();
  for(var i = lastHeader; i < vals.length; i++){
    var row = vals[i];
    if(row[ifCol-1] == ifVal){
      if(andCol && regEx){
        if(!regEx.test(row[andCol-1])){
          continue;
        }
      }
      sheet.getRange(i+1,thisCol).setValue(thisVal);
    }
  }
}
function updateAgeAtStart(sheet,row,birthday){
  if(!birthday){
    birthday = getDateCell(sheet,row,birthdayCol);
  }
  if(!birthday){return;}
  var ageCell = sheet.getRange(row,ageCol);
  if(birthday == "18+"){
    ageCell.setValue("18+");
    updateWillBe18(sheet,row,"18+");
    return;
  }
  var start = getDateCell(sheet,row,startDateCol);
  if(!start){return;}
  start = new Date(start).valueOf();
  var today = new Date().valueOf();
  if(start < today){start = today;}
  var dif = (start-(new Date(birthday).valueOf()))/31540000000;
  ageCell.setValue(dif);
  ageCell.setNumberFormat("0.00");
  tryTryAgain(function(){ //*4/9/20 KJE #240 try again on service error
    updateWillBe18(sheet,row,dif);
  });
}
function getDateCell(sheet,row,col){
  try{
    var range = sheet.getRange(row,col);
    var ret = range.getDisplayValue();
    if(isNaN(ret)){
      return ret;
    }
    if(!(/\d{1,2}\/\d{1,2}\/\d{4}/.test(ret))){
      ret = addMillenium(ret);
      range.setValue(ret);
      range.setNumberFormat("M/d/yyyy");
    }
    return ret;
  }catch(e){
    var list = getServerErrorList();
    for(var i = 0; i < list.length; i++){
      if((e.message).indexOf(list[i])){
        Utilities.sleep(2000);
        return getDateCell(sheet,row,col);
      }
    }
    if(e.message == "The coordinates of the range are outside the dimensions of the sheet."){
      return "";
    }else{
      throw e;
    }
  }
}
/**
 * Updates the "Turned/Will Turn 18?" column for an individual employee based on the appropriate info and updates subsequent columns like whether they need a background check.
 * 
 * 10/2/2019 KJE #166 add isDCF
 * 3/20/24 KJE #404 previously, we assumed staff didn't need a background check if they had a previous contract. Now, consider whether staff turned 18 since their last contract, in which case they would need a background check.
 * 
 * @param {object} sheet The Hired sheet.
 * @param {Number} row The row of the staff in question.
 * @param {object} ageAtStart How old the employee will be on their start date, either a number or "18+".
 * @param {boolean} isDCF Whether this employee needs a DCF background check or not.
 * @author Kyle Egerdal 
 */
function updateWillBe18(sheet,row,ageAtStart,isDCF){
  var recentStart,start;
  if(!ageAtStart){
    ageAtStart = sheet.getRange(row,ageCol).getDisplayValue();
  }
  if(!ageAtStart){return;}
  if(ageAtStart == "18+"){ageAtStart = 100;}
  recentStart = getDateCell(sheet,row,recEmpCol);
  if(!(/\d{1,2}\/\d{1,2}\/\d{4}/.test(recentStart))){
    recentStart = "";
  }
  if(ageAtStart >= 18){
    sheet.getRange(row,t18Col).setValue("x");
    if(!recentStart){
      paperwork18Cols(sheet,row,true,ageAtStart,null,isDCF); //*10/2/2019 KJE #166 isDCF
      return;
    }
  }else{
    sheet.getRange(row,t18Col).setValue("");
    if(!recentStart){
      paperwork18Cols(sheet,row,false,ageAtStart,null,isDCF); //*10/2/2019 KJE #166 isDCF
      return;
    }
  }
  start = getDateCell(sheet,row,startDateCol);
  if(!start){return;}
  start = new Date(start).valueOf();
  var today = new Date().valueOf();
  if(start < today){start = today;}
  var recentStartAsDate = new Date(recentStart);
  var howLongAgo = (start-(recentStartAsDate.valueOf()))/31560000000; //how long ago the last start date was, in years
  if(ageAtStart >= 18){ //*3/20/24 KJE #404 needs paperwork if turned 18 since last contract started
    var birthday = sheet.getRange(row,birthdayCol).getDisplayValue();
    if(birthday != "18+"){
      var birthdayAsDate = new Date(birthday);
      birthdayAsDate.setFullYear(birthdayAsDate.getFullYear()+18);
      if(birthdayAsDate > recentStartAsDate){ //*3/20/24 KJE #404 needs paperwork if turned 18 since last contract started
        paperwork18Cols(sheet,row,true,ageAtStart,howLongAgo,isDCF); //*10/2/2019 KJE #166 isDCF
        return;
      }
    }
    if(howLongAgo && howLongAgo >= 1){ //will turn 18 or is 18 and last worked more than a year ago
      paperwork18Cols(sheet,row,true,ageAtStart,howLongAgo,isDCF); //*10/2/2019 KJE #166 isDCF
      return;
    }
  }
  paperwork18Cols(sheet,row,false,ageAtStart,howLongAgo,isDCF); //*10/2/2019 KJE #166 isDCF
}
/**
 * Updates paperwork columns for an employee based on the provided information.
 * 
 * 10/2/2019 KJE #166 isDCF
 * 9/7/2022 YMC #426 removed needs18 so that 18+ and 18- end up with the same setup
 * 
 * @param {object} sheet The hired sheet.
 * @param {Number} row The row the employee in question is in on the hired sheet.
 * @param {boolean} needs18 Whether this employee needs a background check.
 * @param {Number} ageAtStart How old this employee will be on their start date.
 * @param {Number} howLongAgo How long ago this employee's previous contract terminated.
 * @param {boolean} isDCF Whether this employee needs a DCF background check. 
 * @author Kyle Egerdal
 */
function paperwork18Cols(sheet,row,needs18,ageAtStart,howLongAgo,isDCF){
  //*3/21/24 KJE #519 set Approval column since it works the same as the others, even though it isn't dependent on age
  setNACols(sheet,row,"N/A",[approvalCol]);
  //Paperwork only needed if !howLongAgo or howLongAgo >= 1 year
  var set = (!howLongAgo || howLongAgo >= 1) ? "" : "N/A";
  setNACols(sheet,row,set,[W4Col,W4Col+1,WT4Col,WT4Col+1,DDCol,DDCol+1,I9Col,I9Col+1]);
  //Background check - if over 18 and !howLongAgo or howLongAgo >= 1
  var set = needs18 ? "" : "N/A"; //*3/20/2024 KJE #404 howLongAgo is already accounted for in updateWillBe18. This used to be (needs18 && (!howLongAgo || howLongAgo >= 1)) ? "" : "N/A"
  setNACols(sheet,row,set,[BCCol,BCCol+1]);
  var set = needs18 ? "N/A" : ""; //*3/20/24 KJE #404 if no background check required, have PA double check
  setNACols(sheet,row,set,[BCCheckCol]);
  //*10/2/2019 KJE #166 DCF -- if over 18, !howLongAgo, and isDCF
  if(isDCF !== false){ //don't bother calculating if we know it's false
    var set = (!howLongAgo && (isDCF == true || getDCF(sheet.getRange(row,SARLib.munisCodeCol).getDisplayValue()))) ? "" : "N/A"; //9/7/22 YMC #426 removed needs18 so that 18+ and 18- end up with the same setup
  }else{
    var set = "N/A";
  }
  setNACols(sheet,row,set,[SARLib.DCFCol]);
  //TB test - over 18 and last employment more than three months ago
  var set = (!howLongAgo || (howLongAgo > 0.25)) ? "" : "N/A"; //0.25 = three months *4/1/20 KJE #226 all ages need TBRA
  setNACols(sheet,row,set,[TBCol,TBCol+1]);
  //wp < 16
  var set = (ageAtStart < 16 && (!howLongAgo || howLongAgo >= 1)) ? "" : "N/A";
  setNACols(sheet,row,set,[WPCol]);
  //photo - if !howlongAgo
  var set = (howLongAgo) ? "N/A" : "";
  setNACols(sheet,row,set,[photoCol]);
}
function setNACols(sheet,row,set,cols){
  tryTryAgain(function(){ //*1/24/22 KJE replace "Service error" catch with tryTryAgain
    for(var i = 0; i < cols.length; i++){
      var range = sheet.getRange(row,cols[i]);
      var val = range.getValue();
      if(!set && val == "N/A"){ //if it shouldn't be N/A but it is, clear it
        range.setValue(set);
      }else if(!val){ //if it should be N/A set it to N/A, but don't override if something is there
        range.setValue(set);
      }
    }
  });
}
//*10/2/2019 KJE #166
function getDCF(munis){
  if(!munis){
    return false;
  }
  getMUNISSheet();
  var vals = munisSheet.getDataRange().getDisplayValues();
  for(var i = munisSheetLastHeader; i < vals.length; i++){
    if(Number(vals[i][SARLib.munisSheetCodeCol-1]) == Number(munis)){ //*10/29/2019 KJE match with leading zeros
      return (vals[i][SARLib.munisSheetDCFCol-1] == "Yes");
    }
  }
  return false;
}
function replaceListAssignments(vals,row,col,event){
  var checkThis,checkThisCol;
  var checkThisToo,checkThisColToo;
  var replaceThisCol;
  switch(col){
    case listSheetSiteORCol:
      checkThis = vals[row-1][listSheetSiteCol-1];
      checkThisCol = siteCol-1;
      replaceThisCol = directorCol-1;
      break;
    case listSheetProgORCol: //special handling because of override
      checkThis = vals[row-1][listSheetProgCol-1];
      checkThisCol = programCol-1;
      checkThisToo = "^((?!\@\#\@"; //start of string + not (negative lookahead) + something obscure (for sites without overrides)
      for(var i = listSheetLastHeader; i < vals.length; i++){
        if(vals[i][listSheetSiteORCol-1]){ //site has an override
          checkThisToo = checkThisToo + "|" + vals[i][listSheetSiteCol-1]; //add (not) the site
        }
      }
      checkThisToo = new RegExp(checkThisToo + ").)*$","");
      //regex now looks like this: ^((?!A|B|C).)*$
      checkThisColToo = siteCol-1;
      replaceThisCol = directorCol-1;
      break;
    case listSheetPARetCol:
      checkThis = vals[row-1][listSheetProgCol-1];
      checkThisCol = programCol-1;
      checkThisToo = /^\d{6}$/;
      checkThisColToo = bNumCol-1;
      replaceThisCol = PACol-1;
      break;
    case listSheetPANewCol:
      checkThis = vals[row-1][listSheetProgCol-1];
      checkThisCol = programCol-1;
      checkThisToo = new RegExp("^" + newEmpKey + "$","");
      checkThisColToo = bNumCol-1
      replaceThisCol = PACol-1;
      break;
    default: //safety
      closePleaseWait();
      return;
  }
  var hiredSheet = getHiredSheet()
  var oldRange = hiredSheet.getRange(hiredSheetLastHeader+1,1,hiredSheet.getLastRow(),lastKeepCol);
  var hiredVals = oldRange.getValues();
  var newVal = event.value;
  var replacedOne = false;
  for(var i = hiredSheetLastHeader; i < hiredVals.length; i++){
    if(hiredVals[i][checkThisCol] == checkThis){
      if(!checkThisColToo){
        hiredVals[i][replaceThisCol] = newVal;
        replacedOne = true;
        continue;
      }else if(checkThisToo.test(hiredVals[i][checkThisColToo])){
        hiredVals[i][replaceThisCol] = newVal;
        replacedOne = true;
      }
    }
  }
  if(replacedOne){
    oldRange.setValues(hiredVals);
  }
}
function setSup(sheet,row,site,prog){
  try{
  if(!site){
    site = sheet.getRange(row,siteCol).getValue();
  }
  if(!prog){
    prog = sheet.getRange(row,programCol).getValue();
  }
  var vals = getListSheet().getDataRange().getValues();
  var toSet;
  if(site){
    for(var i = listSheetLastHeader; i < vals.length; i++){
      if(vals[i][listSheetSiteCol-1] == site){ //site sets payrollApvsCol
        sheet.getRange(row,payrollApvsCol).setValue(vals[i][listSheetSiteORCol-1]);
        break;
      }
    }
  }
  if(prog){
    for(var i = listSheetLastHeader; i < vals.length; i++){
      if(vals[i][listSheetProgCol-1] == prog){ //program sets directorCol
        sheet.getRange(row,directorCol).setValue(vals[i][listSheetProgORCol-1]);
        break;
      }
    }
  }
  return vals;
  }catch(e){
    if((e.message).indexOf("Service") != -1){
      Utilities.sleep(2000);
      return setSup(sheet,row,site,prog);
    }else{
      throw e;
    }
  }
}
function setPA(sheet,row,program,listSheetVals,explicit,phantom){
  var PA = explicit;
  if(!PA){
    var getCol = (sheet.getRange(row,bNumCol).getValue() == newEmpKey) ? true : false;
    if(!listSheetVals){
      listSheetVals = getListSheet().getDataRange().getValues();
    }
    var PA = SARLib.getPA(program,getCol,listSheetVals,true); //*3/22/24 KJE #429 add "true" to specify PA for paperwork; function now defaults to PA for payroll
  }
  if(PA){
    if(phantom){
      return PA;
    }
    var range = sheet.getRange(row,PACol);
    if(range.getValue() != PA){
      sheet.getRange(row,PACol).setValue(PA);
      //notifyPADialogue(PA,sheet.getRange(row,idCol).getValue()); *5/13/2020 KJE #241 consolidate PA notifications
    }
  }
  if(phantom){
    return null;
  }
}
//*5/13/2020 KJE #241 delete notifyPA functions---consolidated within AtNight

//
function settedProgram(sheet,row,program,jcCell){
  var munisData = getMUNISSheet().getDataRange();
  var munisVals = munisData.getValues();
  var munisColors = munisData.getBackgrounds();
  var jobs = [];
  var alls = 0;
//KJE 9/10/2018 - only show inclusion services jobs if inclusion services program was chosen
//  for(var i = SARLib.munisSheetLastHeader; i < SARLib.munisSheetLastAll; i++){
//    var val = munisVals[i][munisSheetJobCol-1];
//    if(jobs.indexOf(val) == -1){
//      jobs.push(val);
//      alls++;
//    }
//  }
  var start = false; //MUNIS is a long list and it's alphabetical. Once we start matching, we can stop looking as soon as we stop matching.
  //That is, if the list is A A B B C C and we're looking for a B, once we get to the first B, we can stop as soon as we hit not B.
  var i = SARLib.munisSheetLastHeader; //*KJE 8/5/2019 deprecate munisSheetLastAll
  //!All
  for(; i < munisVals.length; i++){
    if(munisVals[i][SARLib.munisSheetSiteCol-1] != "!All"){
       //i++; *10/8/2019 KJE #173 show first job after !All
       break;
    }
    //*4/13/2021 KJE #302 show all !Alls regardingless of if double-listed with a non-!All site
    if(munisVals[i][munisSheetProgramCol-1] == program){
      if(munisIgnoreColors.indexOf(munisColors[i][munisSheetProgramCol-1]) != -1){continue};
      var job = munisVals[i][munisSheetJobCol-1];
      if(jobs.indexOf(job) == -1){
        jobs.push(job);
      }
    }
  }
  if(i >= munisVals.length){
    i = SARLib.munisSheetLastHeader;
  }
  //Not !All
  for(; i < munisVals.length; i++){
    if(munisVals[i][munisSheetProgramCol-1] == program){ //cols are 1-indexed, values are 0-indexed
      start = true;
      if(munisIgnoreColors.indexOf(munisColors[i][munisSheetProgramCol-1]) != -1){continue};
      var job = munisVals[i][munisSheetJobCol-1];
      if(jobs.indexOf(job) == -1){
        jobs.push(job);
      }
    }else{
      if(start){
        break;
      }
    }
  }
  var jcRule = SpreadsheetApp.newDataValidation().requireValueInList(jobs, true).setAllowInvalid(false).build();
  jcCell.setDataValidation(jcRule);
  if(jobs.length == alls+1){
    var job = jobs[alls];
    jcCell.setValue(job);
    settedJob(sheet,row,job);
  }
  var listVals = setSup(sheet,row,null,program);
  setPA(sheet,row,program,listVals);
}
/**
 * 1/24/22 MDH #371 Ignore grey background cells
 */
function settedJob(sheet,row,job,allowGrandfather){
  try{
    //get MUNIS code
    var munisCell = sheet.getRange(row,munisCodeCol,1,1);
    var program = sheet.getRange(row,programCol,1,1).getValue();
    var start = false; //see notes on this variable above
    var MUNIS = getMUNISSheet().getDataRange().getValues();
    for(var i = 0; i < MUNIS.length; i++){
      if(MUNIS[i][munisSheetProgramCol-1] == program){ //cols are 1-indexed, values are 0-indexed
        start = true;
        if(MUNIS[i][munisSheetJobCol-1] == job){
          munisCell.setValue(MUNIS[i][munisSheetCodeCol-1]);
          //*10/2/2019 KJE #166 DCF
          var isDCF = (MUNIS[i][SARLib.munisSheetDCFCol-1] == "Yes");
          var curDCF = (sheet.getRange(row,SARLib.DCFCol,1,1).getValue() == "N/A");
          if(isDCF && curDCF || (!isDCF && !curDCF)){
            if(!isDCF){
              setNACols(sheet,row,"N/A",[SARLib.DCFCol]);
            }else{
              updateWillBe18(sheet,row,null,isDCF);
            }
          }
          break;
        }
      }else{
        if(start){
          break;
        }
      }
    }
    //get pay rates
    var PRSheet = getPayRateSheet().getDataRange();
    var PRs = PRSheet.getDisplayValues();
    var colors = PRSheet.getBackgrounds(); //some pay rates are ignored by color
    try{
      var jCode = jCodeREx.exec(job)[0];
    }catch(e){
      if((e.message).indexOf("Cannot read property \"0\" from null.") != -1){
        getUi().alert("Job code didn't match M### format, so couldn't find pay rates.");
        return;
      }else{
        throw e;
      }
    }
    var rates = [];
    for(var i = 0; i < PRs.length; i++){
      if(PRs[i][PRSheetCodeCol-1].indexOf(jCode) != -1){
        for(var j = PRSheetPRStartCol-1; j < PRSheet.getNumColumns(); j++){
          if(payRateIgnoreColors.indexOf(colors[i][j]) != -1){continue}; //* 1/24/22 MDH #371 Ignore grey background cells
          if(!allowGrandfather && grandfatherPayRateColors.indexOf(colors[i][j]) != -1){continue};
          var rate = PRs[i][j];
          if(rate){
            if(rate.indexOf(".") == -1){ //column is formatted as currency so we need to format the selections that way too
              rate = rate + ".00";
            }
            rates.push(rate);
          }
        }
        //found the right job, no need to keep looking
        break;
      }
    }
    var showDropDown = (rates.length == 1) ? false : true;
    var payRule = SpreadsheetApp.newDataValidation().requireValueInList(rates, showDropDown).setAllowInvalid(false).build();
    var payCell = sheet.getRange(row,payRateCol,1,1);
    payCell.setDataValidation(payRule);
    payCell.setValue(rates[0]); //choose first pay rate by default
  }catch(e){
    if((e.message).indexOf("Service error") != -1){
      Utilities.sleep(2000);
      settedJob(sheet,row,job,allowGrandfather);
    }else{
      throw e;
    }
  }
}
function grandfatherPayRate(){
  makePleaseWait();
  getSht();
  getUi();
  if(sheet.getName() != hiredSheetName){
    closePleaseWait();
    ui.alert("This tool can only be used on the \"" + hiredSheetName + "\" sheet.");
    return;
  }
  var rows = getActiveRows();
  if(!rows){
    closePleaseWait();
    ui.alert("First select a row with an employee on it.");
    return;
  }
  if(rows.length > 1){
    closePleaseWait();
    ui.alert("Select only ONE row with an employee on it to use this tool.");
    return;
  }
  var row = rows[0];
  settedJob(sheet,row,sheet.getRange(row,jobCodeCol,1,1).getValue(),true);
  closePleaseWait();
}
function clearedJobCodeCol(sheet,row){
  sheet.getRange(row,munisCodeCol,1,1).setValue("");
  var payRateCell = sheet.getRange(row,payRateCol,1,1);
  payRateCell.clearDataValidations();
  payRateCell.setValue(chooseJob);
}
function getNextEntryID(){
  try{
    //prevent adding entries with the same ID#
    var lock = LockService.getScriptLock();
    var haveLock = false, tries = 0;
    while(!haveLock && tries < 5){  //try 5 times
      haveLock = lock.tryLock(1000); //try for 1 second. 5 tries * 1 second each = 5 seconds.
      tries++;
    }
  }catch(e){
    if(e.message == "There are too many LockService operations against the same script."){
      Utilities.sleep(2000);
      return getNextEntryID();
    }else{
      throw e;
    }
  }
  if(!haveLock){
    throw new Error("Error assigning TCID. Please try again in 10 seconds.");
  }
  SpreadsheetApp.flush(); //*5/10/21 KJE attempt to fix duplicate TCIDs
  var range = getSS().getSheetByName("Seed").getDataRange();
  var ret = range.getValue();
  range.setValue(ret+1);
  try{
    lock.releaseLock();
  }catch(e){
    if(e.message == "There are too many LockService operations against the same script."){
      Utilities.sleep(2000);
      lock.releaseLock();
    }else{
      throw e;
    }
  }
  return ret;
}
/**
 * Uses a toast with the message "Working..." or a custom message, if passed.
 * 
 * @param {String} [customMessage] A custom message to display in the toast. The default is "Working..."
 * @author Kyle Egerdal
 */
function makePleaseWait(customMessage){
  try{
    SpreadsheetApp.getActiveSpreadsheet().toast(customMessage ? customMessage : "Working...",pleaseWaitTitle,40);
  }catch(e){
    if((e.message).indexOf("Service") != -1){
      Utilities.sleep(1000);
      makePleaseWait("SEDbot is a little overworked right now. Sorry about that!");
      return;
    }else{
      throw e;
    }
  }
  return;
  //No HTML for installable triggers for users other than the one who installed the trigger
  //var htmlOutput = HtmlService.createHtmlOutput("").setHeight(pleaseWaitHeight);
  //SpreadsheetApp.getUi().showModalDialog(htmlOutput, pleaseWaitTitle);
}
/**
 * Toasts "Done!"
 * 
 * @author Kyle Egerdal
 */
function closePleaseWait(){
  try{
    SpreadsheetApp.getActiveSpreadsheet().toast("Done!",pleaseWaitTitle,3);
  }catch(e){
    if((e.message).indexOf("Service") != -1){
      Utilities.sleep(1000);
      makePleaseWait();
      return;
    }else{
      throw e;
    }
  }
  return;
  //var htmlOutput = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>").setHeight(pleaseWaitHeight);
  //SpreadsheetApp.getUi().showModalDialog(htmlOutput, pleaseWaitTitle);
}
function duplicateLine(){
  makePleaseWait();
  getUi();
  getSht();
  var name = sheet.getName();
  if(name != hiredSheetName && name != activeSheetName && name != archiveSheetName){
    ui.alert("This option can only be used on the " + hiredSheetName + ", " + activeSheetName + ", or " + archiveSheetName + " sheet.");
    return;
  }
  var rows = getActiveRows();
  if(rows.length == 0){
    ui.alert("Oops! Select a row with staff on it and try again.");
    closePleaseWait();
    return;
  }
  if(rows.length > 5){
    var areSure = ui.alert("You're trying to duplicate " + rows.length + " employees.\n\nAre you sure?\n\n(to multi-select with a filter view, be sure to hold ctrl -- don't use shift or click and drag)", ui.ButtonSet.YES_NO_CANCEL)
    if(areSure == ui.Button.CANCEL || areSure == ui.Button.NO || !areSure){
      closePleaseWait();
      return;
    }
  }
  var newRow;
  for(var i = 0; i < rows.length; i++){
    var row = rows[i];
    try{
      newRow = duplicateThis(sheet.getRange(row,allowPasteMin,1,allowPasteMax-allowPasteMin).getDisplayValues()[0]);
    }catch(e){
      if((e.message).indexOf("Must be") != -1 || (e.message).indexOf("data validation") != -1){
        //ignore
      }else{
        throw e;
      }
    }
  }
  var hiredSheet = getHiredSheet();
  sheet.getParent().setActiveSheet(hiredSheet).setActiveRange(getHiredSheet().getRange(newRow,lastNameCol));
  closePleaseWait();
}
function duplicateThis(data){
  var newData = [];
  for(var i = 0; i < allowPasteMin-1; i++){
    newData.push("");
  }
  for(var i = 0; i < data.length; i++){
    newData.push(String(data[i]).trim());
  }
  var sheet = getHiredSheet();
  //only permitted to add to first empty row AFTER DATA. Not first empty row (if there is data after it).
  var newRowNum = sheet.getLastRow()+1;
  var maxRows = sheet.getMaxRows();
  if(newRowNum == maxRows){ //last row is blank as it's supposed to be
    sheet.getRange(newRowNum,1,1,newData.length).setValues([newData]);
  }else{ //last row isn't blank. shouldn't be this way, but handle it.
    sheet.appendRow(newData);
    //setClearedFormat(sheet,maxRows+1,sheet.getRange(maxRows+1,1,1,sheet.getMaxColumns())); *2/1/2019 KJE let myOnEdit handle cleared formatting
  }
  sheet.appendRow([""]);
  try{ //*6/5/2019 KJE
    fillNewRow(sheet,newRowNum); //*2/1/2019 KJE let myOnEdit handle fillNewRow *6/4/2019 KJE myOnEdit doesn't handle multiples
  }catch(e){
    if((e.message).indexOf("protected") != -1){
      //do nothing
    }else{
      throw e;
    }
  }
  return newRowNum;
}
/**
 * 2/21/22 MDH tryTryAgain
 */
function setRowPermissions(){
  var sheet = tryTryAgain(function(){ //*11/3/21 KJE tryTryAgain
    return getSht();
  });
  //Google is remiss and editing a protection while not on the sheet that protection protects will change it to protect the active sheet
  if(sheet.getName() != hiredSheetName){
    return;
  }
  //make sure there's at least one empty row at the bottom
  var maxRows = sheet.getMaxRows();
  if(maxRows <= sheet.getLastRow()){
    tryTryAgain(function(){ //*8/28/31 KJE try again on time out
      sheet.appendRow([""]);
      setClearedFormat(sheet,maxRows+1,sheet.getRange(maxRows+1,1,1,sheet.getMaxColumns()));
    });
  }
  //set permissions for all but the last row (no permissions on last row allows SDs to append/add rows)
  var lastProtectedRow = tryTryAgain(function(){ return sheet.getLastRow(); }); // *2/21/22 MDH tryTryAgain
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++){
    var protection = protections[i];
    var rangeA1 = protection.getRange().getA1Notation();
    if(rangeA1.match(/\d+/)[0] <= hiredSheetLastHeader){ //ignore header ranges
      continue;
    }
    var colSets = rangeA1.match(/[A-Z]+/ig);
    tryTryAgain(function(){ //*11/3 KJE tryTryAgain wrapper
      protection.setRange(ss.getRange(colSets[0] + (hiredSheetLastHeader + 1) + ":" + colSets[1] + lastProtectedRow));
    });
  }
}
var nameREx = /\<|\@|\>/;
function empCanBeCleared(row,sheet){
  var range = sheet.getRange(row,1,1,sheet.getMaxColumns());
  var vals = range.getValues()[0];
  var missing = [];
  if(!vals[idCol-1]){
    missing.push("\n");
    missing.push("- A TCID ID (should have been added automatically -- contact administrator to get one. DO NOT make one up yourself!)");
  }
  var firstName = vals[firstNameCol-1];
  if(!firstName){
    missing.push("\n");
    missing.push("- A first name.");
  }else if(nameREx.test(firstName)){
    missing.push("\n");
    missing.push("- A first name that does not include an e-mail address (<, @, or > symbols).");
  }
  var lastName = vals[lastNameCol-1];
  if(!lastName){
    missing.push("\n");
    missing.push("- A last name.");
  }else if(nameREx.test(lastName)){
    missing.push("\n");
    missing.push("- A last name that does not include an e-mail address (<, @, or > symbols).");
  }
  var bNum = vals[bNumCol-1];
  if((!/\d{6}/.test(bNum)) && (!/V\d+/.test(bNum))){
    missing.push("\n");
    missing.push("- A 6-digit B-number.");
  }
  if(!vals[emailCol-1]){
    missing.push("\n");
    missing.push("- An e-mail address.");
  }
  if(!vals[siteCol-1]){
    missing.push("\n");
    missing.push("- A site.");
  }
  if(!vals[programCol-1]){
    missing.push("\n");
    missing.push("- A program.");
  }
  if(!vals[jobCodeCol-1]){
    missing.push("\n");
    missing.push("- A job.");
  }
  if(!vals[payRateCol-1]){
    missing.push("\n");
    missing.push("- A pay rate.");
  }
  if(!vals[supCol-1]){
    missing.push("\n");
    missing.push("- A director/supervisor.");
  }
  if(!vals[PACol-1]){
    missing.push("\n");
    missing.push("- A PA.");
  }
  return missing.toString().replace(/,/g,"");
}
/*
* First entry for a row. Set data validations, formatting, and formulas
*/
function fillNewRow(sheet,row){
  var range = sheet.getRange(row, 1, 1, sheet.getMaxColumns());
  range.clearFormat();
  setStatusColor("",sheet,row,range);
  var vals = range.getDisplayValues()[0];
  //idCol
  if(!vals[idCol-1]){
    vals[idCol-1] = "#" + getNextEntryID();
    range.getCell(1,idCol).setHorizontalAlignment("center")
  }
  //statusCol
  var statusCell = range.getCell(1,statusCol);
  statusCell.clear();
  statusCell.setDataValidation(SpreadsheetApp.newDataValidation()
                                               .requireValueInList(hiredStatusList, true)
                                               .setAllowInvalid(false).build()).setHorizontalAlignment("left");
  //dateAddedCol
  //if(!vals[dateAddedCol-1]){ sometimes people put last name in the date column by accident. the bottom row isn't protected so they can do that.
    vals[dateAddedCol-1] = Utilities.formatDate(new Date(), tz, "MM/dd/YY");
  range.getCell(1,dateAddedCol).setNumberFormat("M/d/yyyy"); //*11/21/2019 KJE #189 explicitly format
  //}
  //birthdayCol
  range.getCell(1,birthdayCol).setDataValidation(SpreadsheetApp.newDataValidation()
                                                 .requireFormulaSatisfied("=REGEXMATCH(TO_TEXT("+getANotation(birthdayCol)+row+"),\"^\\d{1,2}/\\d{1,2}/\\d{4}|18\\+$\")")
                                                 .setHelpText("Must be a date (MM/DD/YYYY) or \"18+\".")
                                                 .setAllowInvalid(false).build()).setHorizontalAlignment("right");
  //emailCol
  range.getCell(1,emailCol).setDataValidation(SpreadsheetApp.newDataValidation()
                                                 .requireFormulaSatisfied("=REGEXMATCH(TO_TEXT("+getANotation(emailCol)+row+"),\"^[^ ,]+@[^ ,]+\\.[^ ,]+$\")")
                                                 .setHelpText("Enter a single valid e-mail address.")
                                                 .setAllowInvalid(false).build()).setHorizontalAlignment("left");
  //bNumCol
  range.getCell(1,bNumCol).setDataValidation(SpreadsheetApp.newDataValidation()
                                             .requireFormulaSatisfied("=REGEXMATCH(TO_TEXT("+getANotation(bNumCol)+row+"),\"^NEW$|^\\d{6}$|^VNEW$|^V\\d+$\")")
                                             .setHelpText("Must be \"NEW\", a B# (six digits without a \"B\" in front), or VNEW/V# for vendors.")
                                             .setAllowInvalid(false).build());
  //phNumCol
  range.getCell(1,phCol).setDataValidation(SpreadsheetApp.newDataValidation()
                                           .requireFormulaSatisfied("=REGEXMATCH(TO_TEXT("+getANotation(phCol)+row+"),\"^\\d{3}-\\d{3}-\\d{4}$\")")
                                             .setHelpText("Must be ###-###-####.")
                                             .setAllowInvalid(false).build());
  //txtOkCol
  range.getCell(1,txtOkCol).setDataValidation(SpreadsheetApp.newDataValidation()
                                              .requireValueInList(["Yes","No"])
                                              .setAllowInvalid(false).build()).setHorizontalAlignment("right");
  //siteCol
  getMUNISSheet();
  var validCol = getANotation(munisSheetSiteCol);
  var munisVals = munisSheet.getDataRange().getDisplayValues();
  var lastAll = getMUNISLastAll(munisVals);
  var validRange = munisSheet.getRange(validCol+lastAll+":"+validCol+munisSheet.getLastRow());
  range.getCell(1,siteCol).setDataValidation(SpreadsheetApp.newDataValidation()
                                             .requireValueInRange(validRange, true)
                                             .setAllowInvalid(false).build()).setHorizontalAlignment("left");
  //programCol
  if(!vals[programCol-1]){
    sheet.getRange(row,programCol).clearDataValidations().setHorizontalAlignment("left");
    vals[programCol-1] = chooseSite;
  }
//  var validCol = getANotation(munisSheetProgramCol);
//  var validRange = munisSheet.getRange(validCol+(munisSheetLastHeader+1)+":"+validCol+munisSheetLastRow);
//  range.getCell(1,programCol).setDataValidation(SpreadsheetApp.newDataValidation()
//                                             .requireValueInRange(validRange, true)
//                                             .setAllowInvalid(false).build()).setHorizontalAlignment("left");
  //jobCodeCol
  if(!vals[jobCodeCol-1]){
    sheet.getRange(row,jobCodeCol).clearDataValidations().setHorizontalAlignment("left");
    vals[jobCodeCol-1] = chooseProgram;
  }
  //payRateCol
  if(!vals[payRateCol-1]){
    sheet.getRange(row,payRateCol).clearDataValidations().setHorizontalAlignment("right");
    vals[payRateCol-1] = chooseJob;
  }
  //supCol,payRollApvsCol
  getListSheet();
  var listSheetLastRow = listSheet.getLastRow();
  var validCol = getANotation(listSheetSDCol);
  var lastvalidCol = getANotation(listSheetPACol);
  var validRange = listSheet.getRange(validCol+(listSheetLastHeader+1)+":"+lastvalidCol+listSheetLastRow);
  var supValidation = SpreadsheetApp.newDataValidation().requireValueInRange(validRange, true).setAllowInvalid(false).build();
  range.getCell(1,supCol).setDataValidation(supValidation).setHorizontalAlignment("left");
  range.getCell(1,payrollApvsCol).setDataValidation(supValidation).setHorizontalAlignment("left");
  //startDateCol
  var validDate = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  range.getCell(1,startDateCol).setDataValidation(validDate);
  //endDateCol *10/8/2019 KJE #167 validate
  var validCol = getANotation(SARLib.listSheetEndDatesCol);
  var validRange = listSheet.getRange(validCol+(listSheetLastHeader+1)+":"+validCol+listSheetLastRow);
  var endDateValidation = SpreadsheetApp.newDataValidation().requireValueInRange(validRange, true).setAllowInvalid(false).build();
  range.getCell(1,endDateCol).setDataValidation(endDateValidation);
  //emailSentCol
  range.getCell(1,emailSentCol).setDataValidation(validDate);
  //remindCol
  range.getCell(1,SARLib.remindCol).setDataValidation(SpreadsheetApp.newDataValidation()
                                                      .requireValueInList(SARLib.remindList)
                                                     .setAllowInvalid(false).build());
  vals[SARLib.remindCol-1] = SARLib.remindIf;
  //PACol
  var validCol = getANotation(listSheetPACol);
  var validRange = listSheet.getRange(validCol+(listSheetLastHeader+1)+":"+validCol+listSheetLastRow);
  range.getCell(1,PACol).setDataValidation(SpreadsheetApp.newDataValidation()
                                             .requireValueInRange(validRange, true)
                                             .setAllowInvalid(false).build());
  range.setValues([vals]);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  setBorders(sheet,row,range);
}
function getMUNISLastAll(munisVals){
  var lastAll = SARLib.munisSheetLastHeader; //*8/5/19 KJE deprecate SARLib.munisSheetLastAll
  for(; lastAll < munisVals.length; lastAll++){
    var munisRow = munisVals[lastAll];
    if(munisRow[SARLib.munisSheetSiteCol-1] != "!All"){
      lastAll++;
      break;
    }
  }
  if(lastAll >= munisVals.length){
    lastAll = SARLib.munisSheetLastHeader+1;
  }
  return lastAll;
}
function setBorders(sheet,row,range){
  //adffe9
  range.setBorder(true,true,true,true,true,true,"#cccccc",null);
  range.getCell(1,lastKeepCol).setBorder(null,null,null,true,null,null); //set right border of lastKeepCol
  range.getCell(1,notesCol).setBorder(null,true,null,null,null,null); //set left border of notesCol
  sheet.getRange(row,ageCol,1,lastEditCol-ageCol+1).setBorder(null,null,null,null,true,null,"#999999",null);//set horiz borders of paperwork cols
}
function setSDbackground(sheet,row,customStartCol,customEndCol,adtlCol){
  sheet.getRange(row,idCol).setHorizontalAlignment("center");
  if(customStartCol && customEndCol){
    sheet.getRange(row,customStartCol,1,customEndCol-customStartCol+1).setBackground(sdBackgroundColor);
  }else{
    sheet.getRange(row,lastNameCol,1,endDateCol-lastNameCol+1).setBackground(sdBackgroundColor);
  }
  if(adtlCol){
    sheet.getRange(row,adtlCol,1,1).setBackground(sdBackgroundColor);
  }
}
/**
 * 11/9/2021 MDH #336 replaced functionality
 */
function setStatusColor(status,sheet,row,range,customStartCol,customEndCol,adtlCol){
  if(!status){
    range.setBackground("#ffffff").setFontColor("#000000").setFontLine("none");
    setSDbackground(sheet,row,customStartCol,customEndCol,adtlCol);
  }else if(status == cleared){
    range.setBackground("#ffff00").setFontColor("#000000").setFontLine("none");
  }else if(status == quit){
    range.setBackground("#545454").setFontColor("#ffffff").setFontLine("line-through");  //background gray, font white
  }else if(status == missingInfo){
    range.setBackground("#FF0000").setFontColor("#000000").setFontLine("none");  //background red, font black
  }else if(status == deleteMe){
    range.setBackground("#ffffff").setFontColor("#999999").setFontLine("none");
  }
  else if(status == replaced){ //*11/9/2021 MDH #336 replaced functionality
    range.setBackground("#d0e0e3").setFontColor("#999999").setFontLine("none");
  }
}
function setClearedFormat(sheet,row,range){
  range.setBackground("#ffffff").setFontColor("#000000").setFontLine("none");
  setSDbackground(sheet,row);
}
function updateVortex(TCID){
  getUi();
  var choice = ui.alert(getUpdateVortexString(), ui.ButtonSet.YES_NO)
  switch(choice){
    case ui.Button.NO:
      return;
    case ui.Button.YES:

  }
}
function getUpdateVortexString(){
  var ret = "You have changed information that determines how this employee's time is logged.\n\n";
  ret = ret + "If they have already logged time that should have been logged with the updated information, that time needs to be updated.\n\n";
  ret = ret + "Do you want to update their already logged time?";
  return ret;
}
function canLogTime(){
  makePleaseWait();
  getSht();
  getUi();
  if(sheet.getName() != hiredSheetName){
    closePleaseWait();
    ui.alert("This tool can only be used on the \"" + hiredSheetName + "\" sheet.");
    return;
  }
  var rows = getActiveRows();
  if(!rows){
    closePleaseWait();
    ui.alert("First select a row with an employee on it.");
    return;
  }
  if(rows.length > 1){
    closePleaseWait();
    ui.alert("Select only ONE row with an employee on it to use this tool.");
    return;
  }
  var row = sheet.getRange(rows[0],1,1,sheet.getMaxColumns()).getDisplayValues()[0];
  var errors = [];
  var start = row[startDateCol-1];
  if(!start){
    errors.push("\n- They don't have a start date.");
  }else if(new Date(addMillenium(start)).valueOf() > (new Date()).valueOf()){
    errors.push("\n- Their start date hasn't passed yet (" + start + ").");
  }
  if(!row[idCol-1]){
    errors.push("\n- They don't have a TCID. Contact " + SARLib.getErrorEmails(env) + " to have one assigned.");
  }
  if(!row[firstNameCol-1]){
    errors.push("\n- They don't have a first name.");
  }
  if(!row[lastNameCol-1]){
    errors.push("\n- They don't have a last name.");
  }
  if(!row[jobCodeCol-1]){
    errors.push("\n- They don't have a job.");
  }
  if(!row[supCol-1]){
    errors.push("\n- They don't have a supervisor listed.");
  }
  if(!row[PACol-1]){
    errors.push("\n- They don't have a PA listed.");
  }
  if(errors.length == 0){
    var site = row[siteCol-1];
    if(!site){
      site = "(no site listed)";
    }
    var program = row[programCol-1];
    if(!program){
      program = "(no program listed)";
    }
    closePleaseWait();
    ui.alert("This employee should be able to log time under the site \"" + site + "\" or program \"" + program + "\".\n\nIf you have PERSONALLY checked and their name does not appear, please write " + SARLib.getErrorEmails(env) + ".");
  }else{
    closePleaseWait();
    ui.alert("This employee cannot log time for the following reasons:" + errors.toString().replace(/,/g,""));
  }
}
// Years stated as ## become 19##, not 20##. If the date ends in /##, replace with /20##.
function addMillenium(dt){
  if((yearEx).exec(dt)){
    var slash = dt.lastIndexOf("/");
    return dt.substring(0,slash+1) + "20" + dt.substring(slash+1);
  }
  return dt;
}
//function forceUpdatePAs(vals){
//  getSS();
//  var old = ss.getActiveSheet();
//  ss.setActiveSheet(getHiredSheet());
//  updatePAs(vals);
//  ss.setActiveSheet(old);
//}
function test032224_1448(){
  updatePAs(getListSheetVals());
}
function updatePAs(vals){
  if(!vals){
    var vals = getListSheet().getDataRange().getDisplayValues();
  }
  var pas = getPAs(vals);
  var paPermitted = pas.concat(getMasters(vals));
  paPermitted = unique(paPermitted);
  ss.setActiveSheet(getHiredSheet());//Google is dumb and editing a protection while not on the sheet that protection protects will change it to protect the active sheet.
  setEditors(paPermitted,getPAHiredProtections(),true);
  ss.setActiveSheet(getActiveSheet());//see above
  setEditors(paPermitted,getPAActiveProtections(),true);
  updateDrivePermissions(paPermitted,DriveApp.getFileById("1viAQ-3o5jMQaWYXo8suHqSOxzEhDZLHZuMgND0juYF0"),true); //to do
  updateDrivePermissions(paPermitted,getEmpFolder(),true); //folder with all paperwork
  updateDrivePermissions(pas,getCannedResponsesFolder(),true,true); //canned responses (needed to generate agreements) - read only
  updateDrivePermissions(paPermitted,getPACSVsFolder(),true,true); //PA CSV folder - read only
  updateDrivePermissions(paPermitted,getCoverSheetFolder(),true); //PA cover sheet folder
  var thisSS = getSS();
  updateDrivePermissions(getSDs(vals).concat(paPermitted),thisSS); //this spreadsheet!
  updateDrivePermissions(["cwoodward@madison.k12.wi.us","mcanicoba@madison.k12.wi.us","rjstern@madison.k12.wi.us"],thisSS,true,true); //chad - read only *6/3/24 KJE #530 add Ramon
  //*1/5/23 KJE add mcanicoba per email from Julia
  updateAdmins(vals,paPermitted);
}
function updatePAProtectionsOnly(){
  var vals = getListSheet().getDataRange().getDisplayValues();
  var pas = getPAs(vals);
  var paPermitted = pas.concat(getMasters(vals));
  paPermitted = unique(paPermitted);
  ss.setActiveSheet(getHiredSheet());//Google is dumb and editing a protection while not on the sheet that protection protects will change it to protect the active sheet.
  setEditors(paPermitted,getPAHiredProtections(),true);
  ss.setActiveSheet(getActiveSheet());//see above
  setEditors(paPermitted,getPAActiveProtections(),true);
}
function getPAHiredProtections(){
  var ret = [];
  var protections = getHiredSheet().getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var statusCL = getANotation(statusCol);
  var lastKeepCL = getANotation(lastKeepCol+1);
  var CCCL = getANotation(CCCol);
  for (var i = 0; i < protections.length; i++){
    var protection = protections[i];
    var rangeA1 = protection.getRange().getA1Notation();
    var colSets = rangeA1.match(/[A-Z]+/ig);
    if(colSets[0] == statusCL && colSets[1] == statusCL){
      ret.push(protection);
    }else if(colSets[0] == lastKeepCL || colSets[0] == CCCL){
      ret.push(protection);
    }
  }
  return ret;
}
/**
 * 11/19/2021 MDH #354 remove status access for PAs 
 */
function getPAActiveProtections(){
  var ret = [];
  var protections = getActiveSheet().getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var statusCL = getANotation(statusCol);
  var lastNameCL = getANotation(lastNameCol);
  var siteCL = getANotation(siteCol);
  var quitCL = getANotation(quitDateCol);
  for (var i = 0; i < protections.length; i++){
    var protection = protections[i];
    var rangeA1 = protection.getRange().getA1Notation();
    var colSets = rangeA1.match(/[A-Z]+/ig);
    if(colSets[0] == statusCL && colSets[1] == statusCL){
      // ret.push(protection); //*11/19/2021 MDH #354 remove status access for PAs 
    }else if(colSets[0] == lastNameCL && colSets[1] == siteCL){
      ret.push(protection);
    }else if(colSets[0] == quitCL && colSets[1] == quitCL){
      ret.push(protection);
    }
  }
  return ret;
}
function updateMasters(vals){
  var masters = getMasters(vals);
  getSS();
  var old = ss.getActiveSheet();
  var sheets = ss.getSheets();
  //ss protections
  for(var i = 0; i < sheets.length; i++){
    var sheet = sheets[i];
    ss.setActiveSheet(sheet);
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).concat(sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET));
    setEditors(masters,protections);
  }
  //admins get same access as PAs
  updatePAs(vals);
  ss.setActiveSheet(old);
  //admins can edit canned responses
  updateDrivePermissions(masters,getCannedResponsesFolder(),true); //canned responses (needed to generate agreements)
  updateDrivePermissions(masters,DriveApp.getFolderById("1ZfBu0MkD4qMFl8GPOvlXMRYAwWGb4FZa"),true); //tb test dist folder
  updateDrivePermissions(masters,DriveApp.getFolderById("1A-CTUY0M8GhaXtxrX-HwvZO89Pi0ziBL"),true); //mass upload folder
  updateDrivePermissions(masters,DriveApp.getFolderById("1kiAC0smt9VViYu1xIuelXLjxik-zTaEP"),true); //*11/12/2019 KJE #186 blank forms
  updateDrivePermissions(masters,DriveApp.getFileById("0B4o44ub0NVA4RHhlQVJsWmlJcWgzTXl5VGZQS1JnQzlfWEZn"),true,true); //zebra scanner config
  //*7/16/23 KJE #445 hotfix use properties to store masters
  SARLib.storeMasters(env,masters);
}
function updateSDs(vals){
  updateDrivePermissions(getSDs(vals).concat(getPAs(vals)).concat(getMasters(vals)),DriveApp.getFileById(getSS().getId())); //this spreadsheet
}
function updateAdmins(vals,pasAndMasters){
  if(env != "PRD"){ //SARLib.PRDApvSheets only applies to PRD
    return;
  }
  var admins = getAdmins(vals);
  if(!pasAndMasters){
    pasAndMasters = getPAs(vals).concat(getMasters(vals));
  }
  admins = admins.concat(pasAndMasters);
  updateDrivePermissions(admins,DriveApp.getFolderById(SARLib.PRDApvSheets),null,true);
}
function setEditors(list,protections,isUnique){
  var emails = [];
  for(var i = 0; i < list.length; i++){
    var person = list[i];
    var lt = person.indexOf("<");
    if(lt != -1){
      person = person.substring(person.indexOf("<")+1,person.indexOf(">"));
    }
    person = person.toUpperCase();
    emails.push(person);
  }
  list = emails;
  if(!isUnique){list = unique(list);}
  for(var i = 0; i < protections.length; i++){
    var protection = protections[i];
    var editors = protection.getEditors();
    for(var j = 0; j < list.length; j++){
      var person = list[j];
      if(!person){continue;}
      var foundPerson = false;
      for(var k = 0; k < editors.length; k++){
        if(person.indexOf(editors[k].toString().toUpperCase()) != -1){
          foundPerson = true;
          editors.splice(k,1);
          break;
        }
      }
      if(!foundPerson){
        protection.addEditor(person);
      }
    }
    for(var j = 0; j < editors.length; j++){
      protection.removeEditor(editors[j]);
    }
  }
}
function unique(arr){
  var u = {}; ret = [];
  for(var i = 0; i < arr.length; i++){
    var elem = arr[i];
    if(u[elem] === undefined){
      u[elem] = 1;
      ret.push(elem);
    }
  }
  return ret;
}
function extractUniqueEmails(list){
  var emails = [];
  for(var i = 0; i < list.length; i++){
    var person = list[i];
    var lt = person.indexOf("<");
    if(lt != -1){
      person = person.substring(person.indexOf("<")+1,person.indexOf(">"));
    }
    person = person.toUpperCase();
    if(emails.indexOf(person) == -1){
      emails.push(person);
    }
  }
  return emails;
}
/**
 * KJE 7/25/2019 #138 don't remove editors when setting readOnly. Deprecate isUnique via extractUniqueEmails
 * 1/16/21 MDH hotfix #257 catch error when removing deprecated user
 */
function updateDrivePermissions(list,driveItem,isUnique,readOnly,forceUpdate){
  list = extractUniqueEmails(list);
  var currentAccess,editors;
  if(readOnly){
    currentAccess = driveItem.getViewers();
    editors = driveItem.getEditors();
    editors.forEach(function(el,dx,ar){ar[dx] = el.getEmail().toUpperCase()});
  }else{
    currentAccess = driveItem.getEditors();
  }
  currentAccess.push(driveItem.getOwner());
  for(var i = 0; i < list.length; i++){
    var person = list[i];
    if(!person){continue;}
    var foundPerson = false;
    for(var j = 0; j < currentAccess.length; j++){
      if(person.indexOf(currentAccess[j].getEmail().toUpperCase()) != -1){
        foundPerson = true;
        currentAccess.splice(j,1);
        break;
      }
    }
    if(!foundPerson){
      if(env == "PRD" || forceUpdate){ //notifications go out on this, so only do it when we really mean it
        person = person.toLowerCase();
        if(readOnly){
          Logger.log(person);
          driveItem.addViewer(person);
        }else{
          driveItem.addEditor(person);
        }
      }else{
        Logger.log("adding " + (readOnly ? "viewer" : "editor") + " " + person);
      }
    }
  }
  for(var i = 0; i < currentAccess.length; i++){
    if(env == "PRD" || forceUpdate){
      try{ //*6/10/20 KJE #257 catch error when removing deprecated user
        if(readOnly){
          if(editors.indexOf(currentAccess[i].getEmail().toUpperCase()) != -1){
            continue;
          }
          driveItem.removeViewer(currentAccess[i]);
        }else{
          driveItem.removeEditor(currentAccess[i]); //KJE this was removeViewer 1/7/19. ???
        }
      }catch(e){ //*6/10/20 KJE #257 catch error when removing deprecated user
        if((e.message).indexOf("No such user") != -1){ //*1/16/21 MDH hotfix #257 catch error when removing deprecated user
          //do nothing
        }else{
          throw e;
        }
      }
    }else{
      Logger.log("removing " + (readOnly ? "viewer" : "editor") + " " + currentAccess[i].getEmail());
    }
  }
}
function getSDs(vals){
  if(!vals){
    vals = getListSheetVals();
  }
  return getList(vals,SARLib.listSheetSDCol);
}
function getFullSDs(vals){
  return getList(vals,SARLib.listSheetSDCol,true);
}
function getPAs(vals){
  return getList(vals,listSheetPACol);
}
function getFullPAs(vals){
  return getList(vals,listSheetPACol,true);
}
function getProcessors(vals){
  return getList(vals,listSheetProcessorCol);
}
function getAdmins(vals){
  return getList(vals,SARLib.listSheetSDSupsCol);
}
function getMasters(vals){
  if(!vals){  //*3/11/23 KJE #445
    vals = getListSheetVals();
  }
  try{
    var ret = getList(vals,listSheetMasterCol);
    var owner = getSS().getOwner().getEmail();
    if(ret.toString().indexOf(owner) == -1){
      ret.push(owner);
    }
    return ret;
  }catch(e){
    if((e.message.toUpperCase()).indexOf("SERVICE") != -1){
      Utilities.sleep(2000);
      return getMasters(vals);
    }else{
      throw e;
    }
  }
}
function getList(vals,col,leaveFull){
  var ret = [];
  for(var i = listSheetLastHeader; i < vals.length; i++){
    var person = vals[i][col-1];
    if(!leaveFull){
      var bracket = person.indexOf("<");
      if(bracket != -1){
        person = person.substring(bracket+1,person.indexOf(">"));
      }
    }
    if(person){
      ret.push(person);
    }//else{ //KJE 9/17/2018 prevent not adding people below a break
//      break;
//    }
  }
  return ret;
}
