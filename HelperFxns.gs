//Copyright (c) Kyle Egerdal 2018. All Rights reserved.

/**
 * Returns a list of keys whose inclusion can be used to identify Google server errors.
 * 
 * 5/6/2020 KJE #250 "is missing"
 * 7/7/2021 KJE #416 add "simultaneous invocations"
 * 
 * @return {Array<string>} The list of keys.
 * @author Kyle Egerdal
 */
function getServerErrorList(){
  return ["Service","server","LockService","form data","is missing","simultaneous invocations"]; //*5/6/2020 KJE #250 "is missing" //*7/7/22 KJE #416 add "simultaneous invocations"
}
/**
 * Given a function, calls it. If it throws a server error, catches the error, waits a bit, then tries to call the function again. Repeats until the function is executed successfully or a maximum number of tries is reached. If the latter, throws the error.
 * 
 * The idea being that Google often asks users to "try again soon," so that's what this function does.
 * 
 * 1/27/22 MDH #365 add inBetweenAttempts
 * 
 * @param {function} fx The function to call.
 * @param {number} [iv=500] The time, in ms, the wait between calls. The default is 500.
 * @param {number} [maxTries=3] The maximum number of attempts to make before throwing the error. The default is 3.
 * @param {Array<string>} [handlerList=getServerErrorList()] The list of keys whose inclusion can be used to identify errors that cause another attempt. The default is the list returned by getServerErrorList().
 * @param {number} [tries=0] The number of times the function has already tried. This value is handled by the function. The default is 0.
 * @param {function} inBetweenAttempts This function will be called in between attempts. Use this parameter to "clean up" after a failed attempt. 
 * @return {object} The return value of the function.
 * @author Kyle Egerdal
 */
function tryTryAgain(fx,iv,maxTries,handlerList,tries,inBetweenAttempts){
  try{
    return fx();
  }catch(e){
    if(!iv){
      iv = 500;
    }
    if(!maxTries){
      maxTries = 3; 
    }
    if(!handlerList){
      handlerList = getServerErrorList(); 
    }
    if(!tries){
      tries = 1; 
    }
    if(tries >= maxTries){
      throw e; 
    }
    for(var i = 0; i < handlerList.length; i++){
      if((e.message).indexOf(handlerList[i]) != -1){
        Utilities.sleep(iv);
        if(inBetweenAttempts){ inBetweenAttempts();} //*1/27/22 MDH #365 add inBetweenAttempts
        return tryTryAgain(fx,iv,maxTries,handlerList,tries+1,inBetweenAttempts); //*1/27/22 MDH #365 add inBetweenAttempts
      }
    }
    throw e;
  }
}
var capREx = /\b\w+/g;
//*2/19/20 KJE #217 don't capitalize after "di"
//*2/26/20 KJE #222 capitalize first letter only, IGNORE other letters. Also ignore "de " (with space)
function cap(val){
  if(!val){return val};
  val = val.trim();
  var de = (val.indexOf("de ") == 0);
  //capitalize the first letter only  //*9/12/20 KJE #282 unless it's all caps, then also make the rest lowercase
  val = val.charAt(0).toUpperCase() + (val === val.toUpperCase() ? val.substr(1).toLowerCase() : val.substr(1));
  //uncapitalize "de "
  if(de){
    val = val.replace(/\bDe /,"de ");
  }
  return val;
}            
/*
* Some functions will "try again shortly" if an initial attempt fails.
*
* Those second tries can become third tries, and so on, until no more triggers can be created.
*
* This function "cleans up" those triggers by getting rid of all triggers for a certain function that are clock-based.
*
* So calling this before "trying again shortly" ensures all previous "try again" triggers are deleted first.
*/
function cleanUpTriggers(name){
  try{
    var triggers = ScriptApp.getUserTriggers(getSSByID());
    if(triggers){
      for(var i = 0; i < triggers.length; i++) {
        var trigger = triggers[i];
        if(trigger.getHandlerFunction() != name){continue;}
        if(trigger.getEventType() == ScriptApp.EventType.CLOCK){
          ScriptApp.deleteTrigger(trigger);
        }
      }
    }
  }catch(e){
    if(e.message == "Unable to talk to trigger service" ||
      (e.message).indexOf("Please wait a bit and try again.") != -1){
      Utilities.sleep(5000);
      cleanUpTriggers(name);
    }else{
      throw e;
    }
  }
}

/**
 * Returns the index of the first cell in the sheet's data range containing "value." If no cell contains the value,
 * returns the empty string.
 *
 * value: a string value to look for
 * horizontal: a boolean. If true, searches horizontal first (across the first row, then across the second row, etc.).
 *  If not true, searches vertical first (down the first column, then down the second, etc.).
 *
 * Assumes: range exists and is a rectangle.
 */
function getCellWithValue(value,horizontal){
  getDimensions();
  if(!horizontal){
    for(var i = 0; i < cols; i++){
      for(var j = 0; j < rows; j++){
        if(vals[j][i] == value){
          return [j+1,i+1];
        }
      }
    }
  }else{
    for(var i = 0; i<rows; i++){
      for(var j = 0; j<cols; j++){
        if(vals[i][j] == value){
          return [i+1,j+1];
        }
      }
    }
  }
  return [-1,-1]
}
function getANotation(column){
  if(column <= 26){
    return String.fromCharCode(column+64);
  }else if(column > 702){
    throw new Error("Can only convert columns of 702 or less."); 
  }
  var remainder = (column % 26);
  var additor = (remainder == 0) ? 1 : 0;
  var firstLetter = String.fromCharCode(Math.floor(column/26) + 64 - additor);
  var secondLetter = String.fromCharCode(remainder + 64 + (26*additor));
  return firstLetter + secondLetter;
}
/**
 * Takes a coordinate as an array and returns the A1 notation of the coordinate as a string.
 *
 * For instance, an input of [0,1] would return "A2".
 *
 * Works only up to columns of 702 (ZZ).
 */
function getA1Notation(coordinate){
  var row = coordinate[0];
  var column = coordinate[1];
  return getANotation(column)+row;
}

function getIDfromCell(row,column){
  return getIDfromUrl(getCellUrl(row,column)); 
}

/*
 * For a cell containing the HYPERLINK formula, returns the hyperlink. Else, returns the empty string.
 */
function getCellUrl(row,column){
  return getUrl((getFormulas())[row][column]);
}
function getUrl(string){
  if(!string){return "";}
  if(string.indexOf("HYPERLINK") == -1){return "";}
  var start=string.indexOf("\"")+1;
  return string.substring(start,string.indexOf("\"",start));
}

var urlREx = /[-\w]{25,}/
function getIDfromUrl(url){
  var ret = urlREx.exec(url);
  if(ret){
    return ret[0];
  }else{
    return ""; 
  }
}
function getIDFromHyperlink(string){
 return getIDfromUrl(getUrl(string));
}
/**
 * 1/10/2022 MDH #350 forceNew formulas
 */
function getFolder(rowNum){
  getFormulas(true); //*1/10/2022 MDH #350 forceNew formulas
  //*5/21/24 hotfix KJE catch "Cannot read properties of undefined" error 
  var folderUrl;
  try{
    folderUrl = getUrl(formulas[rowNum][folderCol-1]);
  }catch(e){
    if(e.message.includes("Cannot read properties of undefined")){
      //do nothing, folderUrl will be null and a new folder will get created in the else statement below
    }else{
      throw e;
    }
  }
  if(folderUrl || folderUrl === undefined){
    var idFromUrl = getIDfromUrl(folderUrl);
    var folder = DriveApp.getFolderById(idFromUrl);
    return folder;
  }else{
    var newFolder = DriveApp.createFolder(vals[rowNum][lastNameCol-1] + ", " + vals[rowNum][firstNameCol -1]);
    range.getCell(rowNum+1,folderCol).setValue("=HYPERLINK(\"" + newFolder.getUrl() + "\",\"x\")");
    newFolder.moveTo(getEmpFolder());
    // getEmpFolder().addFolder(newFolder);
    // DriveApp.removeFolder(newFolder); //remove from root
    return newFolder;
  }
}
function newText(html,text){
  newText(html,text,0); 
}
function newText(html,text,newLines){
  html = html + text;
  for(var i = 0; i < newLines; i++){
    html = newLine(html); 
  }
  return html;
}
function newLine(html){
  return html + "<br/>";
}
function newLines(html,num){
  for(var i = 0; i < num; i++){
    html = newLine(html); 
  }
  return html;
}
function getActiveRows(){
  var rows = [];
  getSht();
  var ranges = sheet.getActiveRangeList().getRanges();
  for(var i = 0; i < ranges.length; i++){
    var range = ranges[i];
    var row = range.getRow();
    var max = row + range.getNumRows();
    for(var j = row; j < max; j++){
      if(j <= hiredSheetLastHeader){ //emp lines only
        continue; 
      }
      rows.push(j);
    }
  }
  return rows;
}

function colLetter(number){
  return String.fromCharCode(Number(number)+1+64);
}