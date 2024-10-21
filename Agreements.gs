//Copyright (c) Kyle Egerdal 2018. All Rights reserved.
//Comment Here
var templateID = "1QmJ9Fyh7caOyq1jk7j7JFT1U5ElrxCE0aHC1PDFYEpI";
var agreementFolder = "https://drive.google.com/drive/folders/1_uL-lNOIHBL2rW_nX_d_UeNZIVn-llSr";
var timeLimit = 5.5*60*1000; //5 mins 30 seconds

/**
 * Makes one or more employment agreements depending on how the user wants to specify the staff. They can either select or multi-select, this run this function over their selection, or they can run this function and input a range of row numbers to create agreements for staff in those rows.
 * 
 * 8/27/20 KJE #280 move tryTryAgain so unused doc can be deleted
 * 
 * @author Kyle Egerdal
 */
function makeAgreements(){
  try{
    var timer = String(new Date().valueOf()); //keep as a string to prevent converting to scientific
    getSht();
    getUi();
    
    if(sheet.getName() != hiredSheetName){
      ui.alert("This option can only be used on the \"Hired\" sheet.");
      return;
    }
    
    //get row numbers
    var useSelected = ui.alert("Use currently selected rows? (\"No\" to enter rows manually)", ui.ButtonSet.YES_NO_CANCEL)
    if(useSelected == ui.Button.CANCEL){return;}
    var rows = [];
    if(useSelected == ui.Button.YES){
      makePleaseWait();
      rows = getActiveRows();
    }else if(useSelected == ui.Button.NO){
      var row1 = ui.prompt("From which row number?","START", ui.ButtonSet.OK_CANCEL);
      if(row1.getSelectedButton() == ui.Button.CANCEL){
        closePleaseWait();
        return;
      }
      var start = row1.getResponseText();
      if(isNaN(start)){
        ui.alert("Input must be a number. Please try again.", ui.ButtonSet.OK);
        closePleaseWait();
        return;
      }
      var row2 = ui.prompt("To which row number?","END \n Use the same number as the start row for single agreement.", ui.ButtonSet.OK_CANCEL);
      if(row1.getSelectedButton() == ui.Button.CANCEL){
        closePleaseWait();
        return;
      }
      var end = row2.getResponseText();
      if(isNaN(end)){
        ui.alert("Input must be a number. Please try again.", ui.ButtonSet.OK); 
        closePleaseWait();
        return;
      }
      if(end < start){ui.alert("\"From\" row number must be lower than \"to\" row number. Please try again.", ui.ButtonSet.OK); return;}
      makePleaseWait();
      for(var i = start; i < end+1; i++){
        rows.push(i); 
      }
    }else{//unknown button
      return; 
    }
    var makeEmails = ui.alert("Draft new/returning hire e-mails to those employees?\n\n(\"No\" to create agreements only)", ui.ButtonSet.YES_NO_CANCEL)
    if(makeEmails == ui.Button.CANCEL){return;}
    if(makeEmails == ui.Button.YES){
      makePleaseWait();
      makeEmails = true;
    }else if(makeEmails == ui.Button.NO){
      makePleaseWait();
      makeEmails = false;
    }else{//unknown button
      return; 
    }
    
    //get today
    var today = Utilities.formatDate(new Date(), "CST", "MM-dd-YY");
    //set vals
    getDispVals();
    
    //make agreements
    var rowNum; //save for error tracking
    var errCt = 0;
    //check if any staff are assigned to another PA
    var userEmail = getUser().getEmail();
    for (var i = 0; i < rows.length; i++){
      if(vals[rows[i]-1][PACol-1].indexOf(userEmail) == -1){
        var resp = ui.alert("At least one of the selected staff is assigned to another PA. Do you still want to create agreement(s) for them?",ui.ButtonSet.YES_NO);
        if(resp == ui.Button.NO){
          closePleaseWait();
          return;
        }else{
          break; 
        }
      }
    }
    for (var i = 0; i < rows.length; i++){
      rowNum = rows[i];
      if(String(new Date().valueOf()) - timer > timeLimit){
        ui.alert("Timed out on row " + rowNum + ". Start again from there.");
        return;
      }
      var row = vals[rowNum-1]; //vals is o-indexed
      try{
        var EA = makeOneAgreement(row,today,rowNum);
      }catch(e){
        if(e.message == "Cannot read property \"3\" from undefined."){ //blank row
          ui.alert("Please try again without selecting any blank rows.");
          return;
        }else if((e.message).indexOf("Conversion") != -1){ //*4/1/20 KJE #234 notify PA if conversion fails
          ui.alert("The agreement for TCID" + sheet.getRange(rowNum,1).getValue() + " couldn't be converted to a pdf. Please try again." + 
            "\n\nIf this happens multiple trimes, try deleting the agreement from both the folder and the \"x\" from the SED column, then regenerating it.");
          return;
          //*8/27/20 KJE #280 move try again internal so unused doc can be deleted
//        }else if(e.message == "We're sorry, a server error occurred. Please wait a bit and try again."){
//          Utilities.sleep(2000);
//          var EA = makeOneAgreement(row,today,rowNum);
        }else{
          throw e; 
        }
      }
      if(makeEmails){
        try{
          if(newHireEmail(row,EA)){ //draft Email
            sheet.getRange(rowNum,emailSentCol).setValue(new Date()); //set EA sent date
          }
        }catch(e){
          if((e.message).indexOf("Invalid") != -1){
            var willCont = (i == rows.length-1) ? "" : "\n\nContinuing with other staff...";
            ui.alert(e.message + " (row " + rowNum + ").\n\nThis could be in the employee's e-mail column, the site director column, or the CC column.\n\nPlease fix this e-mail address before trying again." + willCont);
            errCt++;
            continue;
          }else{
            throw e; 
          }
        }
      }
    }
    closePleaseWait();
    if(rows.length > errCt){
      if(makeEmails){
        ui.alert("E-mails have been drafted to the chosen employees. Agreements are attached to the e-mails.\n\nCheck your drafts folder.");
      }else{
        ui.alert("Agreements created. Use the links in the Agreement column to access them."); 
      }
    }
  }catch(e){
    if(e.message == "Timed out waiting for user response"){return;}
    SARLib.sendAdminEmail(env, e.message + "\n" + e["stack"] + "\n" + "Row: " + rowNum + "\n" + "TCID: " + sheet.getRange(rowNum,1).getValue());
    ui.alert("Error processing your entry. An administrator has been notified.");
  }
}
/**
 * Makes a single pdf employment agreement.
 * 
 * ASSUMES: sheet
 * 
 * 4/1/20 KJE #228 allow tokens anywhere, not just in the body
 * 8/27/20 KJE #280 try again on service error. Delete file if it was created before trying again
 * 
 * @param {Array<string>} data The row on the SED to make an agreement for, containing string values (fetched with getDisplayValues, not getValues).
 * @param {string} today Today's date, as a string.
 * @param {number} rowNum The 1-indexed row number for this employee in the SED, used to update the link to the file.
 * @return {object} The pdf employment agreement.
 * @author Kyle Egerdal
 */
function makeOneAgreement(data,today,rowNum){
  try{ //*8/27/20 KJE #280 try again on service error
    //make a copy of the template
    var docname = (data[lastNameCol-1]+", "+data[firstNameCol-1]+" Contract Created "+today);
    //make a copy of templateUrl with name docname in folder
    var doc = DocumentApp.openById(DriveApp.getFileById(templateID).makeCopy(docname,getFolder(rowNum-1)).getId()); // *4/12/22 MDH #335 make the copy in the emp folder 
    
    //*4/1/2020 KJE #228 allow tokens anywhere, not just in the body
    var replaceTokens = function(element,today,data){
      element.replaceText("%LASTNAME%", data[lastNameCol-1]);
      element.replaceText("%FIRSTNAME%", data[firstNameCol-1]);
      element.replaceText("%DATE%", today);
      element.replaceText("%START%", data[startDateCol-1]);
      element.replaceText("%END%", data[endDateCol-1]);
      element.replaceText("%JOBCODE%", data[jobCodeCol-1]);
      element.replaceText("%PAYRATE%", data[payRateCol-1]);
      element.replaceText("%TCID%", data[idCol-1]);
      element.replaceText("%BNUMBER%", "B"+data[bNumCol-1]);
      element.replaceText("%POSITIONCODE%", data[munisCodeCol-1]);
      element.replaceText("%SITE%", data[siteCol-1]);
      element.replaceText("%PA%", data[PACol-1]);
      
      //remove unused tokens
      element.replaceText("%JOBCODE2%", "");
      element.replaceText("%POSITIONCODE2%", "");
      element.replaceText("%Training%", "");
      element.replaceText("%PAYRATE2%", "");
    }
    
    //fill it
    var body = doc.getBody();
    var container = body.getParent();
    var children = container.getNumChildren();
    for(var i = 0; i < children; i++){
      var child = container.getChild(i);
      var type = child.getType();
      switch(type){
        case DocumentApp.ElementType.HEADER_SECTION:
          replaceTokens(child.asHeaderSection(),today,data);
          break;
        case DocumentApp.ElementType.BODY_SECTION:
          replaceTokens(child.asBody(),today,data);
          break;
        case DocumentApp.ElementType.FOOTER_SECTION:
          replaceTokens(child.asFooterSection(),today,data);
          break;
      }
    }
    
    //save, convert to pdf.
    doc.saveAndClose();
    var ret = DriveApp.createFile(doc.getAs("application/pdf"));
    ret.setName(docname + ".pdf");
    
    //delete original
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    
    //move to folder and document in SED
    var folder = getFolder(rowNum-1); //HelperFxns.gs
    ret.moveTo(folder);
    // folder.addFile(ret); //add to EMP folder
    // DriveApp.removeFile(ret); //remove from root
    sheet.getRange(rowNum,SACol).setFormula("=HYPERLINK(\"" + ret.getUrl() + "\",\"x\")"); //set EA link
    
    //return pdf
    return ret;
  }catch(e){ //*8/27/20 KJE #280 try again on service error. Delete file if it was created.
    var list = getServerErrorList();
    var onList = false;
    for(var i = 0; i < list.length; i++){
      if((e.message).indexOf(list[i]) != -1){
        onList = true;
        break;
      }
    }
    if(ret){
      ret.setTrashed(true);
    }else if(doc){
      tryTryAgain(function(){
        DriveApp.getFileById(doc.getId()).setTrashed(true);
      });
    }
    if(onList){
      Utilities.sleep(1000);
      return makeOneAgreement(data,today,rowNum);
    }else{
      throw e; //handle it farther up
    }
  }
}
/**
 * Drafts an email to a recent hire containing their employment agreement and an appropriate body depending on what documents they need to turn in.
 * 
 * @param {Array<object>} data The row on the SED to make an email for, containing string values (fetched with getDisplayValues, not getValues).
 * @param {object} EA The employment agreement to attach.
 * @return {boolean} True if successful, false if there was an error.
 * @author Kyle Egerdal
 */
function newHireEmail(data,EA){
  var to = data[emailCol-1];
  if(!to){
    GmailApp.createDraft(null, "New Hire Email for staff with ID" + data[idCol-1], "This staff has no e-mail address listed. Please list an e-mail address in the SAR and try again."); 
    return false;
  }
  var CCs = data[supCol-1] + "," + data[CCCol-1];
  if(CCs == ","){CCs = "";} //sometimes there's no supervisor
  var attachments = [];
  attachments.push(EA);
  var replaceArray = {};
  replaceArray["FIRSTNAME"] = data[firstNameCol-1];
  replaceArray["JOBCODE"] = data[jobCodeCol-1];
  replaceArray["PROGRAM"] = data[programCol-1];
  var sup = data[supCol-1]; //*12/12/2019 KJE #202 add supervisor token
  replaceArray["SUPERVISOR"] = sup.substring(0,sup.indexOf("<")-1);
  var id = data[idCol-1];
  if(data[W4Col-1] == "N/A" && data[WT4Col-1] == "N/A" && data[DDCol-1] == "N/A" && data[I9Col-1] == "N/A" && data[WPCol-1] == "N/A"){
    if(data[BCCol-1] == "N/A"){
      if(data[TBCol-1] == "N/A"){
        makeHtmlEmail(to,CCs,attachments,cannedAgtOnly,replaceArray,id);
        return true;
      }else{
        var age = data[ageCol-1]; //*7/20/20 KJE #268 special email if both TB and u18
        if(!isNaN(age) && age < 18){
          makeHtmlEmail(to,CCs,attachments,cannedAgTBu18,replaceArray,id); //*7/20/20 KJE #268 special email if both TB and u18
          return true;
        }else{
          //attachments.push(DriveApp.getFileById(tbraDocId)); //*12/9/2019 KJE #203 don't attach TBRA doc -- send link instead. Per Julia
          //*12/19/2019 KJE #166 DCF
          makeHtmlEmail(to,CCs,attachments,(data[SARLib.DCFCol-1] == "N/A" ? cannedAgtTBRA : cannedAgTBDCF),replaceArray,id);
          return true;
        }
      }
    }else if(data[t18Col-1]){
      makeHtmlEmail(to,CCs,attachments,cannedNoMinor,replaceArray,id);
      return true;
    }
  }
  //*10/8/2019 KJE #166 DCF
  makeHtmlEmail(to,CCs,attachments,(data[SARLib.DCFCol-1] == "N/A" ? cannedNewHire : cannedNewDCF),replaceArray,id);
  return true;
}
