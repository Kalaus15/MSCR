//Copyright (c) Kyle Egerdal 2018. All Rights reserved.
/**
 * 2/28/22 MDH #386 use atNightTrigger 
 */
function setUpTriggers(){
  ScriptApp.newTrigger("atNightTrigger").timeBased().everyDays(1).atHour(0).create();  //*2/28/22 MDH #386 use atNightTrigger 
  var ss = getSSByID();
  ScriptApp.newTrigger("myOnEdit").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("myOnChange").forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger("clearFoldersUtil").timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(17).create();
  if(env == "TST"){
    setProp(EIDProp,String(new Date().valueOf()).substring(5,10));
    getNextEntryID();
    setProp(TSTIDProp,SpreadsheetApp.getActive().getId());
  }
}
function takeDownTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for(var i = 0; i < triggers.length; i++){
    var trigger = triggers[i];
    if(trigger.getHandlerFunction() == "sendReminderEmails"){
      continue;
    }
    ScriptApp.deleteTrigger(trigger);
  }
}
