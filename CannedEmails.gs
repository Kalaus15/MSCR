//Copyright (c) Kyle Egerdal 2018. All Rights reserved.

var cannedNewHire = "1nO4CB4NqoVXQ9JWgA9dQ5Ga78chypViLOoJivv2j_4I";
//*10/8/2019 KJE #166 DCF
var cannedNewDCF  = "1PJCIUed15rk1x2LgMrNiYH0Zy6azGfScAOXCkcHr1uQ";
var cannedAgtOnly = "1M8lyKpoAaAY_KooGsu4LtjV46-rIHTK2BT6m7YoTCCU";
var cannedAgtTBRA = "1yXTpArFCLhvX956DbrX--LPsNMgoET4Sv4Z-bUjod58";
var cannedAgTBu18 = "1uboiImwtg6BWj8TInlXl5XRFQ6Ty1DlEeLxDQM3llWg"; //*7/20/20 KJE #268
var cannedAgTBDCF = "1gXM5CRj8Vht-SBgeaajWXOcdaCK0F_-3SPk6tDx19Ns"; //*12/19/2019 KJE #166
var cannedNoMinor = "1bcArIhi20RXouF3YPNW-kCjrbO1yZvcM2Nwa8f3IAYY";
var cannedRetrRem = "1jl0kCFpPakUXp-ord32LN5Wvt-iZfY6Kymrd_9FzCiw";
var clearedToWork = "17K40zYWkmjlJIckqKaTJCiPkNQK-psatdIy8Dh-sLHE"; //special token: %BNUMBER%
var tbraDocId = "0B4o44ub0NVA4YVRXdnVFdGx2OEllUW9aRkx0MzRXeG5sTmRB";
var cachedEmails = {};
var cachedNames = {};
var cachedSig = "";

function makeHtmlEmail(to,CCs,attachments,docID,replaceArray,id){
  var html = "";
  html = newText(html,getCannedResponse(docID,replaceArray),2);
  html = newText(html,betterGetSignature());
  var subj;
  if(cachedNames[docID]){
    subj = cachedNames[docID]; 
  }else{
    subj = DriveApp.getFileById(docID).getName();
    cachedNames[docID] = subj;
  }
  for(token in replaceArray){
    if(token == "PROGRAM"){
      subj = subj + " " + replaceArray[token];
      break;
    }
  }
  if(id){
    subj = subj + " (ID" + id + ")"; 
  }
  tryTryAgain(function(){
    GmailApp.createDraft(to, subj, "HTML is not enabled in your email client. Sad face!", {
      htmlBody: html,
      cc: CCs,
      attachments: attachments
    });
  });
}
function getCannedResponse(docID,replaceArray){
  var html;
  if(cachedEmails[docID]){
    html = cachedEmails[docID];
  }else{
    var forDriveScope = DriveApp.getStorageUsed(); //needed to get Drive Scope requested
    var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+docID+"&exportFormat=html";
    var param = {
      method      : "get",
      headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions:true,
    };
    html = UrlFetchApp.fetch(url,param).getContentText();
    html = makeCSSInline(html); //docs uses css in the head, but gmail only takes it inline. need to move css inline.
    cachedEmails[docID] = html;
  }
  for(token in replaceArray){
    html = html.replace(new RegExp("%" + token + "%","g"),replaceArray[token]);
  }
  return html;
}
function getSignature(){
  return betterGetSignature();
//  if(!cachedSig){
//    var draft = GmailApp.search("subject:signature label:draft", 0, 1);
//    if(draft.length == 0){return "To include a signature, make a canned response with the subject \"signature.\"";}
//    cachedSig = draft[0].getMessages()[0].getBody();
//  }
//  return cachedSig;
}
/*
* advanced Gmail API must be enabled:
* https://developers.google.com/apps-script/guides/services/advanced
*/
function betterGetSignature(){
  if(!cachedSig){
    cachedSig = Gmail.Users.Settings.SendAs.list("me").sendAs.filter(
      function(account){
        if(account.isDefault){
          return true
        }
      })[0].signature;
  }
  return cachedSig;
}
var classArray=[];
function makeCSSInline(html){
  //handles only .c# clases. no headers (h1), paragraphs (p), spans, etc.
  classArray=[]; //clear in case we are making multiple e-mails
  var headEnd = html.indexOf("</head>");
  //get everything between <head> and </head>, remove quotes
  var head = html.substring(html.indexOf("<head>")+6,headEnd).replace(/"/g,"");
  var pREx = /p{.+?}/;
  var p = pREx.exec(head);
  if(p){
    p = "" + p;
    p = p.substring(2,p.length-1);
  }
  //split on .c# with any positive integer amount of #s
  var regex = /\.c\d{1,}/;
  var classes = head.split(regex);
  //get class info and put in an array index by class num. EG c4{size:small} will put "size:small" in classArray[4]
  var totalLength = 0;
  for(var i = 1; i < classes.length; i++){
    //assume the first string (classes[0]) isn't a class definition
    totalLength = totalLength + classes[i-1].length;
    var cNum = head.substring(totalLength+2,head.indexOf("{",totalLength)); //totallength+2 chops off .c, so get what's between .c and {
    totalLength = totalLength + 2 + cNum.length //add .c and the number of digits in the num
    classArray[cNum] = classes[i].substring(1,classes[i].indexOf("}")); //put what's between .c#{ and } in classArray[#]
  }
  
  //now we have the class definitions, let's put it in the html  
  html = html.substring(headEnd+7,html.indexOf("</html>")); //get everything between <html> and </html>
  var classMatch = /class=\"(c\d{1,} ){0,}(c\d{1,})\"/g
  //matches class="c# c#..." where c#[space] occurs any number of times, even zero times, and c#[no space] occurs after it, exactly once
  html = html.replace(classMatch,replacer); //replace class="c# c#..." with the definitions in classArray[#]
  
  if(p){
    html = html.replace(/<p style="/g,"<p style=\"" + p + ";");
  }
  
  html = html.replace(/;*padding:(\d+pt *)+/,""); //remove padding
  //html = addInheritance(html);
  
  return html;
}

function replacer(match){
  var csOnly = match.substring(7,match.length-1); //class=" has 7 chars, remove the last "
  var cs = csOnly.split(" "); //get each c#
  var ret = "style=\""
  for(var cCount = 0; cCount < cs.length; cCount++){
    ret = ret + classArray[cs[cCount].substring(1)];
    if(ret[ret.length-1] != ";"){ret = ret + ";";}
  }
  return ret+"\"";
}

function addInheritance(html){
  var newHtml = html.substring(0,html.indexOf("style=\""));
  var split = html.split("style=\"");
  var inheritance = {};
  var inheritanceLevel = 0;
  var start = (newHtml.length == 0) ? 0 : 1;
  for(var i = start; i < split.length; i++){
    inheritanceLevel++;
    inheritance[inheritanceLevel] = {};
    var chunk = split[i];
    var endOfStyle = chunk.indexOf("\"");
    declaration = chunk.substring(0,endOfStyle-1);
    var rest = chunk.substring(endOfStyle);
    var declarations = declaration.split(";");
    for(var j = 0; j < declarations.length; j++){
      var item = declarations[j];
      var colon = item.indexOf(":");
      var value = item.substring(colon+1);
      item = item.substring(0,colon);
      inheritance[inheritanceLevel][item] = value;
    }
    chunk = "";
    for(level in inheritance){
      if(level < inheritanceLevel){continue;}
      var thisLevel = inheritance[level];
      for(item in thisLevel){
        if(!(new RegExp("(\"|;) *" + item)).test(chunk)){ //first entry for the item
          chunk = chunk + item + ":" + thisLevel[item] + ";";
        }else{ //already an entry for the item -- override it
          //chunk = chunk.replace(new RegExp("((\"|;)" + item + ":).+?(?:;)"),"$1" + thisLevel[item] + "$2");
        }
      }
    }
    var loseLevels = rest.match(/<\//g);
    if(loseLevels){
      inheritanceLevel = inheritanceLevel-loseLevels.length;
    }
    newHtml = newHtml + "style=\"" + chunk + rest;
  }
  return newHtml;
}