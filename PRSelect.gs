//Copyright (c) Kyle Egerdal 2018. All Rights reserved.

var listRange = "A2:A";
var error = "First select a cell in the Payroll Approvers column."

function payrollSidebar(){
  if(SpreadsheetApp.getActiveRange().getColumn() != payrollApvsCol){
  //for some reason, evaluating this statement within showPayrollError always returns true
    showPayrollError();
    return;
  }
  SpreadsheetApp.getUi().showSidebar(HtmlService.createTemplateFromFile("PRSelectHTML")
                                     .evaluate()
                                     .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                                     .setTitle("Payroll Approvers"));
}
function getOptions(){
  //builds and returns an array of the values in listRange
  return getListSheet()
    .getRange(listSheetLastHeader+1,listSheetSDCol,getListSheet().getLastRow()-listSheetLastHeader-1,1)
    .getDisplayValues()
    .filter(String)
    .reduce(function(a, b){return a.concat(b)});
}
function putOptions(arr){
  var thisRange = SpreadsheetApp.getActiveRange();
  if(thisRange.getColumn() != payrollApvsCol){
    throw new Error(error); 
  }
  SpreadsheetApp.getActiveRange().setValue(arr.join(","));
}
function showPayrollError(){
  if(SpreadsheetApp.getActiveRange().getColumn() != payrollApvsCol){
    var htmlOutput = HtmlService.createHtmlOutput(error).setHeight(50);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Error");
    return true;
  }
  return false;
}