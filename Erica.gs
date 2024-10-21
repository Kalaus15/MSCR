/**
 * 2/25/22 MDH #382 Covid Vaccination
 * 3/20/24 KJE #404 Remove COVID Column
 */
function sendCantWorkEmails(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var sds = {};
  for(var i = 2; i < data.length; i++){
    var row = data[i];
    if(row[programCol-1] != "SREC"){
      continue; 
    }
    if(row[statusCol-1] == "Active"){
      continue; 
    }
    if(row[juliaCol-1]){
      continue; 
    }
    var sd = row[payrollApvsCol-1];
    if(!sd){
      sd = row[supCol-1]; 
    }
    if(!sd){
      continue; 
    }
    if(sds[sd] === undefined){
      sds[sd] = {}; 
    }
    var name = String(row[lastNameCol-1]).trim() + ", " + String(row[firstNameCol-1]).trim();
    var missing = [];
    if(!row[SACol]){
      missing.push("Agreement"); 
    }
    if(!row[W4Col]){
      missing.push("W4"); 
    }
    if(!row[WT4Col]){
      missing.push("WT4"); 
    }
    if(!row[DDCol]){
      missing.push("Direct Deposit"); 
    }
    if(!row[I9Col]){
      missing.push("I9");
    }
    if(!row[TBCol]){
      missing.push("TBRA"); 
    }
    //*3/20/24 KJE #404 Remove COVID Column
    // if(!row[covidVaxCol]){ //* 2/25/22 MDH #382 Covid Vaccination
    //   missing.push("COVID-19 Vaccination"); 
    // }
    if(!row[BCCol]){
      missing.push("Fingerprinting"); 
    }
    if(!row[WPCol-1]){
      missing.push("Work Permit"); 
    }
    if(missing.length == 0){
      missing.push("N/A"); 
    }
    sds[sd][name] = missing;
  }
  for(sd in sds){
    var html = "";
    html = newText(html,"Hello,",2);
    html = newText(html,"You are receiving this e-mail because you have SREC staff that are missing paperwork. This is accurate as of 3:15 PM today, Friday, 6/21.",2);
    html = newText(html,"<b>If the staff below are not cleared to work by the end of the day TODAY, they cannot work on Monday.</b>",2);
    var list = sds[sd];
    for(name in list){
      var missing = list[name];
      html = newText(html,"- " + name + ": " + missing.reduce(function(a, b){return a.concat(", " + b)}),1);
    }
    html = newText(html,"",1);
    html = newText(html,"Thank you,",1);
    html = newText(html,robotName);
    GmailApp.createDraft(sd, "[IMPORTANT] You have SREC staff who are not cleared to work!", "Please view in a browser that renders html",{
      htmlBody: html,
      cc: "Julia T Meyer <jtmeyer@madison.k12.wi.us>, Allison Miller <amiller2@madison.k12.wi.us>, Erica M Pape <empape@madison.k12.wi.us>, Wendy Moran <wlmoran@madison.k12.wi.us>"
    });
  }
}
function tempDeleteAllDrafts(){
  var drafts = GmailApp.getDrafts();
  for(var i = 0; i < drafts.length; i++){
    drafts[i].deleteDraft(); 
  }
}
function ericaReminders(){
  var aptSheet = SpreadsheetApp.openById("1QuZ2wKXghQiD1OlQIZHdgVnCt525lkFxqPLqIWLezok");
  var aptVals = aptSheet.getDataRange().getDisplayValues();
  var hasApt = [];
  for(var i = 1; i < aptVals.length; i++){
    var hasAptEmail = aptVals[i][1];
    if(hasAptEmail){
      hasApt.push(hasAptEmail); 
    }else{
      break; 
    }
  }
  var hired = SpreadsheetApp.openById("1Fi3xZQ-pHDQ-U6RbRIjK24eJrOJjpYQQWnQQpHZ7Qyk").getSheetByName(hiredSheetName);
  var vals = hired.getDataRange().getDisplayValues();
  var cutoffDate = new Date(); //*5/29/19 KJE use two weeks ago
  cutoffDate.setDate(new Date().getDate() - 14);
  var signature = betterGetSignature();
  for(var i = 1; i < vals.length; i++){
    var row = vals[i];
    if(row[SARLib.statusCol-1] == SARLib.cleared){ //nobody cleared
      continue; 
    }
    if(row[SARLib.programCol-1] != "SREC"){ //only SREC
      continue; 
    }
    var email = row[SARLib.emailCol-1];
    if(hasApt.indexOf(email) != -1){ //nobody with apt already
      continue;
    }
    var dateAdded = row[SARLib.dateAddedCol-1];
    var asDate = new Date(addMillenium(dateAdded));
    if(asDate >= cutoffDate){ //nobody added after cutoff date
      continue; 
    }
    var name = row[SARLib.firstNameCol-1].trim();
    var email = row[SARLib.emailCol-1];
    if(GmailApp.search("to:" + email + " \"I hope this email finds you well. We are happy to have you with MSCR this season! We are approaching the MSCR SREC summer season quickly and noticed you still need to complete your paperwork. It is very important that this gets completed prior to starting work this summer. Please see below the previous email sent by\"").length > 0){
      continue; //nobody already reminded 
    }
    var areNew = row[SARLib.bNumCol-1] == "NEW";
    var threads = GmailApp.search("to:" + email + " from:" + (areNew ? "kjegerdal@madison.k12.wi.us" : "wlmoran@madison.k12.wi.us") + " subject:\"" + (areNew ? "New Hire - MSCR SREC" : "Returning Hire -" + "\""));
    if(threads.length == 0){continue;}
    var messages = threads[0].getMessages();
    var message = null;
    for(var j = 0; j < messages.length; j++){
      if(messages[j].getTo() == email){
        message = messages[j];
        break;
      }
    }
    if(message){
      message.createDraftReplyAll("Please view this in a browser that renders html.", {
        htmlBody: oneEricaEmail(name,areNew ? "Kyle Egerdal" : "Wendy Moran",signature) + "<br/><br/>" + message.getBody(),
        replyTo: areNew ? "kjegerdal@madison.k12.wi.us" : "wlmoran@madison.k12.wi.us"
      });
    }
  }
}
function oneEricaEmail(name,pa,signature){
  var html = SARLib.newText("", "Hi " + name + ",", 2);
  html = SARLib.newText(html, "I hope this email finds you well. We are happy to have you with MSCR this season! We are approaching the MSCR SREC summer season quickly and noticed you still need to complete your paperwork. It is very important that this gets completed prior to starting work this summer. Please see below the previous email sent by "+pa+" regarding your paperwork for hire. Please complete the required paperwork as soon as possible. Should you have any questions please don't hesitate to ask.", 2);
  html = SARLib.newText(html, "Best Regards,");
  html = SARLib.newText(html, signature);
  return html;
}