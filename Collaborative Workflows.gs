/*
Preamble

The Kaṇāda Open Source Initiative is an initiative to spread awareness about unrecognized individuals or peoples whose important contributions in 
science, technology, philosophy, language, politics, art, architecture, and culture continue to have an impact to this day.
This is hopefully one of only first such projects which uses the spirit of sharing via the Open Source Software community and turns into a movement 
with the simple goal of helping understand these peoples and their stories using facts and keeping them alive with memories and useful software
Simple rules to follow while contributing with Kaṇāda Open Source Initiative (KOSI)

1. Please use primary or secondary sources of reference such as yours or someone else’s research on the topic, via writings, pictures, or videos
2. Please avoid tertiary sources of reference (i.e. to wikipedia, quora etc.)
3. Please make an effort to publish code which is useful and not just snippets
4. You may use the preamble or inline description for your stories and references e.g. to the Incas and their art or architecture
5. Please include this preamble to your project and if it’s an HTML file you will notice the preamble is enclosed within the <script> element

*/

/*
Collaborative Workflows (Your project name here) is a part of KOSI


Kaṇāda's writings lead us to conclude that he was what we would call a theoretical physicist who validated several thought experiments
His postulations of the laws of motion predate any other modern day scientist.
Therefore just as he set the laws of motion and perhaps inspired modern day physicists and their scientific discoveries,
 so do we set in motion collaborative workflows and making document reviews easier.

Reference -- https://archive.org/details/thevaiasesikasut00kanauoft/mode/2up

With Collaborative Workflows you can prevent accidental sharing and usage of docs by capturing and displaying their workflow state,
so the document collaborator knows, whether the author of the doc is requesting for document approval, 
whether they have all necessary approvals, whether they've made all changes requested, the list of reviewers, and if it's been published
*/

/* कारणाभावात्कार्याभावः ॥१।२।१॥ ~ Kaṇāda, Circa 500 B.C.E.
In the absence of cause there is an absence of effect [motion]

The following functions set up the app within your Add on menu
*/

function doGet(e)
{
  Logger.log(e);
  //onOpen("{user=, source=Spreadsheet, authMode=LIMITED, range=Range}");
  onOpen();
}

function getRightApp()
{
  var uiTypeSpreadsheet = null;
  var uiTypeDocument = null;
  var uiTypeSlides = null;
  var retVal = {};
  
  /* One of the following will succeed */
  
  try {
    uiTypeSpreadsheet = SpreadsheetApp.getUi(); // Function to time.
    Logger.log("This is the spreadsheet getUI output");
    Logger.log(uiTypeSpreadsheet);
    retVal["thisApp"] = SpreadsheetApp;
    retVal["thisAppString"] = "SpreadsheetApp";
  } catch (err) {
    // Logs an ERROR message.
    console.error('SpreadsheetApp.getUi() yielded an error: ' + err + " with UI type " + uiTypeSpreadsheet);
  }

  try {
    uiTypeDocument = DocumentApp.getUi(); // Function to time.
    Logger.log("This is the document getUI output");
    Logger.log(uiTypeDocument);
    retVal["thisApp"]  = DocumentApp;
    retVal["thisAppString"] = "DocumentApp";
  } catch (err) {
    // Logs an ERROR message.
    console.error('DocumentApp.getUi() yielded an error: ' + err + " with UI type " + uiTypeDocument);
  }
  
  try {
    uiTypeSlides = SlidesApp.getUi(); // Function to time.
    Logger.log("This is the document getUI output");
    Logger.log(uiTypeSlides);
    retVal["thisApp"]  = SlidesApp;
    retVal["thisAppString"] = "SlidesApp";
  } catch (err) {
    // Logs an ERROR message.
    console.error('SlidesApp.getUi() yielded an error: ' + err + " with UI type " + uiTypeSlides);
  }
  return retVal;
}

function onOpen(e) {
 
  //var thisApp = this[ e.source + "App" ];
  var uiTypeSpreadsheet = null;
  var uiTypeDocument = null;
  var uiTypeSlides = null;
  var thisApp = null;
  var thisAppString;
  var retVal = getRightApp();
  var lock;
  thisApp = retVal.thisApp;
  thisAppString = retVal.thisAppString;
  
  /* One of the following will succeed */
  try {
    lock = LockService.getScriptLock();
    PropertiesService.getDocumentProperties().setProperty('thisApp', thisAppString);
    lock.releaseLock();
  } catch (err) {
    try {
    PropertiesService.getDocumentProperties().setProperty('thisApp', thisAppString);
    } catch (err) {
      Logger.log("Couldn't set doc property on open");
    }
  }
  
  var ui = thisApp.getUi();
  
  //var ui = SpreadsheetApp.getUi();
  
  ui.createAddonMenu()
  .addItem('Get Started', 'setupFunction')
  .addSeparator()
  .addItem('Being Edited', 'menuEdited')
  .addSeparator()
  .addItem('Under Review', 'menuReview')
  .addSeparator()
  .addItem('Needs Change', 'menuNeedsChange')
  .addSeparator()
  .addItem('Approve', 'menuApprove')
  .addSeparator()
  .addItem('Published', 'menuPublish')
  .addToUi();
  refreshSideBar();
}

function onInstall(e) {
  // This event is used when developing add-ons
  // { authMode: 'LIMITED' or 'FULL' }
  
  // You can run other simple triggers here
  onOpen(e);
}

/*
कार्य्यविरोधि कर्म ॥१।१।१४॥ ~ Kaṇāda, Circa 500 B.C.E.
Action (kārya) is opposed by reaction (karman)

Currently we are only able to set triggers for Spreadsheets due to limitations in the App Script library.
Triggers will change document state to 'Being Edited' from anything else (Under Review, Needs Change, Approved, Published), as soon as it detects any changes to the sheet 
*/

function setupTriggers(thisAppString, ss)
{
  var allTriggers = ScriptApp.getProjectTriggers();
  var lock = LockService.getScriptLock(); 
  for (var i = 0; i < allTriggers.length; i++) { 
    if (allTriggers[i].getHandlerFunction() == 'menuEdited') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }     
  if (thisAppString == "SpreadsheetApp") {
    ScriptApp.newTrigger('menuEdited')  
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  } else if (thisAppString == "DocumentApp") {
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) { 
      if (allTriggers[i].getHandlerFunction() == 'setupFunction') {
        ScriptApp.deleteTrigger(allTriggers[i]);
      }
    }     
    ScriptApp.newTrigger('setupFunction')  
    .forDocument(ss)
    .onOpen()
    .create();
  }  else if (thisAppString == "FormApp") {
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) { 
      if (allTriggers[i].getHandlerFunction() == 'setupFunction') {
        ScriptApp.deleteTrigger(allTriggers[i]);
      }
    }     
    ScriptApp.newTrigger('setupFunction')  
    .forForm(ss)
    .onOpen()
    .create();
  }
  
  PropertiesService.getDocumentProperties().setProperty('triggerState', "set");
  lock.releaseLock();
}

/*
नोदनविशेषाभावान्नोर्ध्वं न तिर्य्यग्गमनम् ॥५।१।८॥ ~ Kaṇāda, Circa 500 B.C.E.
In the absence of a force, there is no upward motion, sideward motion or motion in general.

This function will initialize the document properties and set the initial state of the document if it's not already in motion
*/

function setupFunction(callerFunc) {
  
  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;
  var lock = LockService.getScriptLock();
  var setupCalled = PropertiesService.getDocumentProperties().getProperty('setupCalled');
  lock.releaseLock();
   
  var ui = thisApp.getUi();
  
  var approvalState = null;
  var htmlOutput;
  var listObj = null;
  var localReviewerTempEmails = Array();
  var triggerState = null;
    
  var ss;
  
  if (thisAppString == "SpreadsheetApp") {
    ss = thisApp.getActiveSpreadsheet();
  } else if (thisAppString == "DocumentApp") {
    ss = thisApp.getActiveDocument();
  }  else if (thisAppString == "FormApp") {
    ss = thisApp.getActiveForm();
  } else if (thisAppString == "SlidesApp") {
    ss = thisApp.getActivePresentation();
  }

  
  Logger.log("This App " + thisApp);
  Logger.log("This app String " + thisAppString);
  
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log('Could not obtain lock after 10 seconds.');
    console.log('Could not obtain lock after 10 seconds.');
  }
  
  lock = LockService.getScriptLock();
  triggerState = PropertiesService.getDocumentProperties().getProperty('triggerState');
  lock.releaseLock();
  
  Logger.log("triggerState from myFunction is " + triggerState);

  if(triggerState === null || triggerState === undefined || triggerState === "unset") {
      
    setupTriggers(thisAppString, ss);
  }
  
  lock = LockService.getScriptLock();
  approvalState = PropertiesService.getDocumentProperties().getProperty('approvalState');
  listObj = PropertiesService.getDocumentProperties().getProperty('reviewerListObj');
  lock.releaseLock();
  
  //if (listObj.listObjArr === undefined || listObj.listObjArr === null) {
  Logger.log('List Obj Returned: ' + JSON.stringify(listObj));
    
  if (listObj === undefined || listObj === null) {
    
    localReviewerTempEmails.push({"initialized": "yes"});
   
    lock = LockService.getScriptLock();
   PropertiesService.getDocumentProperties().setProperty('reviewerListObj', JSON.stringify(localReviewerTempEmails));
    lock.releaseLock();
    
    //findUpdateStatus("initialized", "yes", 'reviewerListObj'); //Works
    
    lock = LockService.getScriptLock();
    var localTempListArr = PropertiesService
      .getDocumentProperties().getProperty('reviewerListObj');
    lock.releaseLock();
  
    Logger.log('Stringified Reviewer Email List Returned then parsed and Stringified again before funccall : ' + JSON.stringify(JSON.parse(localTempListArr)));
        
    //localTempListArr = PropertiesService
      //.getDocumentProperties().getProperty('reviewerListObj');
  
    //Logger.log('Stringified Reviewer Email List Returned then parsed and Stringified again after funccall: ' + JSON.stringify(JSON.parse(localTempListArr)));

  }
  
     if (approvalState === undefined || approvalState === null || approvalState == "Being Edited") { /* 3. This is the first time ticket counter is being initialized */
       
       if (callerFunc != "menuEdited") {
         lock = LockService.getScriptLock();
         PropertiesService.getDocumentProperties().setProperty('setupCalled', "yes");
         if (callerFunc == null || callerFunc == undefined)
           PropertiesService.getDocumentProperties().setProperty('approvalState', "Being Edited");
         lock.releaseLock();
         menuEdited();
       }
       
     } else {

       if(null)
          {
          // Put a switch case here
          }
       
       //menuReview();

       refreshSideBar();
     }
  
  lock = LockService.getScriptLock();
  approvalState = PropertiesService
      .getDocumentProperties()
      .getProperty('approvalState');
  lock.releaseLock();
  
  Logger.log('From Setup Approval state is: ' + approvalState);
  
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) { 
    if (allTriggers[i].getHandlerFunction() == 'menuEdited') {
      Logger.log("From Setup Trigger id at index " + i + " is " + allTriggers[i]);
    }
  }
  
  lock = LockService.getScriptLock();
  PropertiesService.getDocumentProperties().setProperty('setupCalled', "yes");
  lock.releaseLock();
 
}

function findUpdateStatus(keyName, newStatus, docKey)
{
  var tempObjArr;
  var foundFlag = 0;
  
  var lock = LockService.getScriptLock();
  tempObjArr = JSON.parse(PropertiesService
      .getDocumentProperties()
       .getProperty(docKey));
      //.getProperty('reviewerListObj'));
  lock.releaseLock();

for (i in Object.keys(tempObjArr))
  {
    Logger.log("Object key: " + Object.keys(tempObjArr[i]).valueOf() + " keyName passed: " + keyName);
    if (Object.keys(tempObjArr[i]).valueOf() == keyName) {
      foundFlag = 1;
      
      for (j in tempObjArr[i])
      {
       Logger.log('key before: ' + Object.keys(tempObjArr[i]).valueOf());
       Logger.log('value before: ' + Object.values(tempObjArr[i]).toString());
       
       tempObjArr[i][j] = newStatus;
       
       Logger.log('key after: ' + Object.keys(tempObjArr[i]).valueOf());
       Logger.log('new value after: ' + Object.values(tempObjArr[i]).toString());
      } 
    }
  }
  
  var editorEmails = getEditorEmails();
  if(!foundFlag && (editorEmails.indexOf(keyName) > -1)) {
    Logger.log("From findUpdateStatus: The keyName passed is an editor to this doc");
      var tempKey = keyName;
      var tempVal = newStatus;
  
      var tempObj = {};
      tempObj[tempKey] = tempVal
      //tempObjArr.push({keyName: tempObj}); //Works but you get {\"keyName\":{\"abcd@gmail.com\":\"Not sent\"}}
      tempObjArr.push(tempObj);
    Logger.log('Pushing key: ' + Object.keys(tempObjArr[i]).valueOf());
    Logger.log('Pushing value: ' + Object.values(tempObjArr[i]).toString());
  }

  lock = LockService.getScriptLock();  
  PropertiesService
      .getDocumentProperties()
      .setProperty(docKey, JSON.stringify(tempObjArr));
  lock.releaseLock();
      //.setProperty('reviewerListObj', JSON.stringify(tempObjArr));
  
  //if(!foundFlag) {
  lock = LockService.getScriptLock();
  Logger.log("This is what we pushed in the doc ", JSON.stringify(PropertiesService
      .getDocumentProperties().getProperty('reviewerListObj')));
  lock.releaseLock();
  //}
}

function sendEmails (reviewerEmail, statusChange, callerFunc)
{
  
  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;
  var emailPreamble;
  var ui = thisApp.getUi();
  var ss;
  var subject;
  var docUrl;
  var retVal = 0;
  var thisApp = this[thisAppString];
  var emailBody;
  
  var lock = LockService.getScriptLock();
  var approvalState = PropertiesService.getDocumentProperties().getProperty('approvalState');
  lock.releaseLock();
  
  if (thisAppString == "SpreadsheetApp") {
    ss = thisApp.getActiveSpreadsheet();
    docUrl = thisApp.getActiveSpreadsheet().getUrl();
  } else if (thisAppString == "DocumentApp") {
    ss = thisApp.getActiveDocument();
    docUrl = thisApp.getActiveDocument().getUrl();
  }  else if (thisAppString == "FormApp") {
    ss = thisApp.getActiveForm();
    docUrl = thisApp.getActiveForm().getUrl();
  } else if (thisAppString == "SlidesApp") {
    ss = thisApp.getActivePresentation();
    docUrl = thisApp.getActivePresentation().getUrl();
  }

  if (statusChange == "Resent" || statusChange == "sent") {
    if (approvalState == 'Published' && callerFunc != "menuPublish") {
      thisApp.getUi() // Or DocumentApp or FormApp.
      .alert('You can\'t Request Review a Published Doc!');
      return retVal;
    } else {
      //thisApp.getUi().alert("This is the status " + statusChange + " and this is the reviewer email " + reviewerEmail);
      findUpdateStatus(reviewerEmail, statusChange, 'reviewerListObj'); // This will do the previous two steps
      
      //if (statusChange == null || statusChange === undefined) {
      emailBody = "Dear, " + reviewerEmail + " \n Please review this document and change the Approval Status of the document to either Approved or Needs Change."
      + "\n" + docUrl + "\n" +"\n Thank you. \n" + Session.getActiveUser().getEmail()
      + "\n\nTo assist in your approval process, you may add the Kaṇāda Approval Workflow Add On";
      
      //menuReview(); //You could turn status to Under Review but do you want to? Yes, absolutely
      retVal = 1;
      emailPreamble = "Request to review doc ";
    }
  } else {
    if (approvalState == 'Published') {
      thisApp.getUi() // Or DocumentApp or FormApp.
      .alert('You can\'t Request change state of a Published Doc!');
      return retVal;
    } else {
      findUpdateStatus(reviewerEmail, statusChange, 'reviewerListObj'); // This will do the previous two steps
      
      emailBody = " \n Please note that the reviewer " + reviewerEmail + " has updated the document status to: " + statusChange
      + "\n" + docUrl + "\n" +"\n Thank you. \n" + Session.getActiveUser().getEmail()
      + "\n\nTo assist in your approval process, you may add the Kaṇāda Approval Workflow Add On";

      retVal = 1;
      emailPreamble = "Approval status change in doc ";
    }
  }
    if (thisAppString == "SpreadsheetApp") {
    
    subject =  emailPreamble + thisApp.getActiveSpreadsheet().getName() + " from " + Session.getActiveUser().getEmail();
    
  } else if (thisAppString == "DocumentApp") {
    
    subject =  emailPreamble + thisApp.getActiveDocument().getName() + " from " + Session.getActiveUser().getEmail();
    
  }  else if (thisAppString == "FormApp") {
    
    subject =  emailPreamble + thisApp.getActiveForm().getName() + " from " + Session.getActiveUser().getEmail();
    
  }  else if (thisAppString == "SlidesApp") {
    
    subject =  emailPreamble + thisApp.getActivePresentation().getName() + " from " + Session.getActiveUser().getEmail();
    
  }
  
  if (statusChange == "Resent" || statusChange == "sent") {
    //MailApp.sendEmail(reviewerEmail, subject, emailBody);
    
    lock = LockService.getScriptLock();
    PropertiesService.getDocumentProperties().setProperty('approvalState', 'Under Review');
    lock.releaseLock();

    refreshSideBar();
  } else {
    var reviewerEmails = getReviewerEmails();
    Logger.log("Here are the reviewer emails: " + JSON.stringify(reviewerEmails));
    for(i in reviewerEmails)
    {
      if (Object.keys(reviewerEmails[i]).toString() == "initialized")
        continue; //Don't get the first value
      //emailBody = "Dear, " + DriveApp.getFileById(ss.getId()).getOwner().getEmail() + emailBody;
      emailBody = "Dear, " + Object.keys(reviewerEmails[i]).toString() + emailBody;
      //MailApp.sendEmail(Object.keys(reviewerEmails[i]).toString(), subject, emailBody);
    }
    refreshSideBar();
  }
    
 //}
  Logger.log("This is the retVal " + retVal);
  /*  if (!retVal)
  ui.alert("Pls enter a valid reviewer Email ID");*/
  return retVal;
}

function showAlert(alertMsg)
{

  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;

  var ui = thisApp.getUi();
  ui.alert(alertMsg);
}

function getEditorEmails ()
{
  var emails;

  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;

  var ss;
  if (thisAppString == "SpreadsheetApp") {
    ss = thisApp.getActiveSpreadsheet();
  } else if (thisAppString == "DocumentApp") {
    ss = thisApp.getActiveDocument();
  }  else if (thisAppString == "FormApp") {
    ss = thisApp.getActiveForm();
  }   else if (thisAppString == "SlidesApp") {
    ss = thisApp.getActivePresentation();
  }

  return ss.getEditors().toString();
}

function getReviewerEmails ()
{
  var emails;
  var tempObjArr;
  var newObjArr = Array();
  tempObjArr = JSON.parse(PropertiesService
      .getDocumentProperties()
      .getProperty('reviewerListObj'));

for (i in Object.keys(tempObjArr))
  {
    Logger.log('From getReviewerEmails key : ' + Object.keys(tempObjArr[i]).valueOf());
    Logger.log('From getReviewerEmails value : ' + Object.values(tempObjArr[i]).toString());
    
    if (Object.keys(tempObjArr[i]).toString() == "initialized") //We don't want the first value, since all it says is initialized, so just skip past it.
      continue;
       newObjArr.push(tempObjArr[i]);
  }
  Logger.log('From getReviewerEmails newObjArr : ' + JSON.stringify(newObjArr));
  return newObjArr;
}

/*
कर्मं कर्मसाध्यं न विद्यते॥१।१।११॥ ~ Kaṇāda, Circa 500 B.C.E.
From motion, [new] motion is not known.
*/

function refreshSideBar()
{
  var htmlOutputFromFile;
  var htmlOutput;
  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;
  
  var ss;
  if (thisAppString == "SpreadsheetApp") {
    ss = thisApp.getActiveSpreadsheet();
  } else if (thisAppString == "DocumentApp") {
    ss = thisApp.getActiveDocument();
  }  else if (thisAppString == "FormApp") {
    ss = thisApp.getActiveForm();
  }  else if (thisAppString == "SlidesApp") {
    ss = thisApp.getActivePresentation();
  }
  
  var ui = thisApp.getUi();
  
  htmlOutputFromFile = HtmlService.createHtmlOutputFromFile('Kaṇāda');
  
  //Concat the html custom output below before the output from the file above
  lock = LockService.getScriptLock();
  htmlOutput = HtmlService.createHtmlOutput('<p>' + PropertiesService.getDocumentProperties().getProperty('approvalState') + '</p>' + htmlOutputFromFile.getContent()).setTitle('Approval state:');
  lock.releaseLock();
  
  ui.showSidebar(htmlOutput);
}

function removeObject(keyName, docKey)
{
  var tempObjArr
  var newObjArr = Array();
  var foundFlag = 0;

  var lock = LockService.getScriptLock();
  tempObjArr = JSON.parse(PropertiesService
      .getDocumentProperties()
       .getProperty(docKey));
      //.getProperty('reviewerListObj'));
  lock.releaseLock();

for (i in Object.keys(tempObjArr))
  {
    Logger.log("Object key: " + Object.keys(tempObjArr[i]).valueOf() + " keyName passed: " + keyName);
    if (Object.keys(tempObjArr[i]).valueOf() == keyName) {
      foundFlag = 1;
      continue;
    }
    newObjArr.push(tempObjArr[i]);    
  }
  
  lock = LockService.getScriptLock();  
  PropertiesService
      .getDocumentProperties()
      .setProperty(docKey, JSON.stringify(newObjArr));
  lock.releaseLock();
      //.setProperty('reviewerListObj', JSON.stringify(tempObjArr));

  lock = LockService.getScriptLock();
  Logger.log("From removeObj: This is what we pushed in the doc ", JSON.stringify(PropertiesService
      .getDocumentProperties().getProperty('reviewerListObj')));
  lock.releaseLock();

}

function addNewItem (e)
{  
  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;

  var ui = thisApp.getUi();
  //ui.alert(JSON.stringify(e));
  return Session.getActiveUser().getEmail();
}

function menuEdited(e) {

  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;

  var ss;
  if (thisAppString == "SpreadsheetApp") {
    ss = thisApp.getActiveSpreadsheet();
  } else if (thisAppString == "DocumentApp") {
    ss = thisApp.getActiveDocument();
  }  else if (thisAppString == "FormApp") {
    ss = thisApp.getActiveForm();
  }  else if (thisAppString == "SlidesApp") {
    ss = thisApp.getActivePresentation();
  }
  
  var ui = thisApp.getUi();
  
  var htmlOutputFromFile;
  var htmlOutput;
  
  var setupCalled = "0";
  
  var lock = LockService.getScriptLock();
  var approvalState = PropertiesService.getDocumentProperties().getProperty('approvalState');
  var triggerState = PropertiesService.getDocumentProperties().getProperty('triggerState');
  setupCalled = PropertiesService.getDocumentProperties().getProperty('setupCalled');
  lock.releaseLock();

  if (setupCalled != "yes") {
    setupFunction("menuEdited"); 
  }
  
  Logger.log("From menuEdited Trigger state is " + triggerState);
  Logger.log("From menuEdited Approval state is " + approvalState);
  
  if (approvalState != "Being Edited")
    
  {
    if ( Session.getActiveUser().getEmail() !=  DriveApp.getFileById(ss.getId()).getOwner().getEmail())
    {
      thisApp.getUi() // Or DocumentApp or FormApp.
      .alert('You can\'t Change State to \"Being Edited\" unless you\'re owner!');
    } else {
      
      if(triggerState == "set") {
        
        lock = LockService.getScriptLock();
        var allTriggers = ScriptApp.getProjectTriggers();
        
       for (var i = 0; i < allTriggers.length; i++) { 
          if (allTriggers[i].getHandlerFunction() == 'menuEdited') {
            ScriptApp.deleteTrigger(allTriggers[i]);
            Logger.log("Trigger id at index " + i + " is " + allTriggers[i]);
          }
        }
        lock.releaseLock();
        
        var reviewerEmails = getReviewerEmails();
        Logger.log("MenuUpdated reviewer emails: " + JSON.stringify(reviewerEmails));
        for(i in reviewerEmails)
        {
          findUpdateStatus(Object.keys(reviewerEmails[i]).toString(), "Resend for review", 'reviewerListObj');
        }
        
        lock = LockService.getScriptLock();
        PropertiesService.getDocumentProperties().setProperty('triggerState', "unset");
        lock.releaseLock();
      }
      
      lock = LockService.getScriptLock();
      PropertiesService.getDocumentProperties().setProperty('approvalState', 'Being Edited');
      //approvalState = PropertiesService.getDocumentProperties().getProperty('approvalState');
      lock.releaseLock();
      
      //thisApp.getUi().alert('Changed state to Being Edited');
      
    }
  }
  
  refreshSideBar();
 
  //ui.createMenu('Approval State: ' + PropertiesService.getDocumentProperties().getProperty('approvalState')).addToUi();
}

/*
Pratyakṣa (प्रत्यक्ष) - Perception ~ Kaṇāda, Circa 500 B.C.E.

When you add a reviewer to the doc their state in the document is automatically changed to 'sent to review' and the document state changes to 'Under Review'
*/

function menuReview(e) {
  
  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;

  var ss;
  if (thisAppString == "SpreadsheetApp") {
    ss = thisApp.getActiveSpreadsheet();
  } else if (thisAppString == "DocumentApp") {
    ss = thisApp.getActiveDocument();
  }  else if (thisAppString == "FormApp") {
    ss = thisApp.getActiveForm();
  }  else if (thisAppString == "SlidesApp") {
    ss = thisApp.getActivePresentation();
  }
  var ui = thisApp.getUi();
  
  var htmlOutputFromFile;
  var htmlOutput;
  
  var setupCalled = "0";  
  var lock = LockService.getScriptLock();
  var approvalState = PropertiesService.getDocumentProperties().getProperty('approvalState');
  var triggerState = PropertiesService.getDocumentProperties().getProperty('triggerState');
  setupCalled = PropertiesService.getDocumentProperties().getProperty('setupCalled');
  lock.releaseLock();
  
  if (setupCalled != "yes") {
    setupFunction("menuEdited"); 
  }
  
  if ( approvalState == 'Published') {
    thisApp.getUi() // Or DocumentApp or FormApp.
    .alert('You can\'t Request Review a Published Doc!');
    return;
  }
  
  if (Session.getActiveUser().getEmail() ==  DriveApp.getFileById(ss.getId()).getOwner().getEmail() ) {
    findUpdateStatus(Session.getActiveUser().getEmail(), "Under Review", 'reviewerListObj');
    
    lock = LockService.getScriptLock();
    PropertiesService.getDocumentProperties().setProperty('approvalState', 'Under Review');
    lock.releaseLock();
    sendEmails(Session.getActiveUser().getEmail(), 'Under Review');
    
  } else {
    
    sendEmails(Session.getActiveUser().getEmail(), 'Under Review');
  }
  
  Logger.log("From menuReview");
  //PropertiesService.getDocumentProperties().setProperty('approvalState', 'Under Review');

  refreshSideBar();
  
  if(triggerState == "unset") {
    
    setupTriggers(thisAppString, ss);
    
  }
  //ui.createMenu('Approval State: ' + PropertiesService.getDocumentProperties().getProperty('approvalState')).addToUi();
}

/*
Anumāna (अनुमान) - Inference ~ Kaṇāda, Circa 500 B.C.E.

A reviewer may change decide the document needsa change and notify the author. This also sends a notification to all the document reviewers
*/

function menuNeedsChange(reviewerEmail) {
  
  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;

  var ss;
  if (thisAppString == "SpreadsheetApp") {
    ss = thisApp.getActiveSpreadsheet();
  } else if (thisAppString == "DocumentApp") {
    ss = thisApp.getActiveDocument();
  }  else if (thisAppString == "FormApp") {
    ss = thisApp.getActiveForm();
  }  else if (thisAppString == "SlidesApp") {
    ss = thisApp.getActivePresentation();
  }
  var ui = thisApp.getUi();

  var htmlOutput;
  
  var setupCalled = "0";  
  var lock = LockService.getScriptLock();
  var approvalState = PropertiesService.getDocumentProperties().getProperty('approvalState');
  var triggerState = PropertiesService.getDocumentProperties().getProperty('triggerState');
  setupCalled = PropertiesService.getDocumentProperties().getProperty('setupCalled');
  lock.releaseLock();
  
  if (setupCalled != "yes") {
    setupFunction("menuEdited"); 
  }
  
  if ( approvalState == 'Published') {
    thisApp.getUi() // Or DocumentApp or FormApp.
    .alert('You can\'t Request Changes In a Published Doc!');
    return null;
  } else {
    
    if (reviewerEmail != null && reviewerEmail != undefined) {
      if (Session.getActiveUser().getEmail() == reviewerEmail) {
        findUpdateStatus(Session.getActiveUser().getEmail(), "Needs Change", 'reviewerListObj');
        
      } else {
        thisApp.getUi() // Or DocumentApp or FormApp.
        .alert('You can\'t change review state for someone else!');
        return null;
      }
    } 
    
    //if ((reviewerEmail == null || reviewerEmail === undefined) && Session.getActiveUser().getEmail() ==  DriveApp.getFileById(ss.getId()).getOwner().getEmail()) { // If I own the doc I can change to any state
    if (Session.getActiveUser().getEmail() ==  DriveApp.getFileById(ss.getId()).getOwner().getEmail() ) {
      findUpdateStatus(Session.getActiveUser().getEmail(), "Needs Change", 'reviewerListObj');
      
      lock = LockService.getScriptLock();
      PropertiesService.getDocumentProperties().setProperty('approvalState', 'Needs Change');
      lock.releaseLock();
      sendEmails(Session.getActiveUser().getEmail(), 'Needs change');
    } else if ((Session.getActiveUser().getEmail() !=  DriveApp.getFileById(ss.getId()).getOwner().getEmail()) && (reviewerEmail == null || reviewerEmail === undefined)) { // This came in via the menu
      //findUpdateStatus(Session.getActiveUser().getEmail(), "Approved", 'reviewerListObj');
      
      //findUpdateStatus(Session.getActiveUser().getEmail(), "Needs change", 'reviewerListObj');
      //PropertiesService.getDocumentProperties().setProperty('approvalState', 'Approved');
      sendEmails(Session.getActiveUser().getEmail(), 'Needs change');
      
      /*thisApp.getUi() // Or DocumentApp or FormApp.
      .alert('You can\'t change state of document from menu if you don\'t own the document!');
      return null;*/
    }

    refreshSideBar();
    
    if(triggerState == "unset") {
      setupTriggers(thisAppString, ss);      
    }
  }
  Logger.log("From Menu Needs Change: Returning Needs change");
  return "Needs change";
}

/*
संयोगाभावे गुरुत्वात् पतनम् ॥५।१।७॥ ~ Kaṇāda, Circa 500 B.C.E.
In the absence of conjunction, gravity [causes objects to] fall.

Either the doc owner (author) can decide to publish the document even while status of other reviewers is not Approved.
In order to publish a doc it needs to be moved to an 'Approved' state by the owner first
*/

function menuApprove(approverEmail) {
  
  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;

  var ss;
  if (thisAppString == "SpreadsheetApp") {
    ss = thisApp.getActiveSpreadsheet();
  } else if (thisAppString == "DocumentApp") {
    ss = thisApp.getActiveDocument();
  } else if (thisAppString == "FormApp") {
    ss = thisApp.getActiveForm();
  } else if (thisAppString == "SlidesApp") {
    ss = thisApp.getActivePresentation();
  }
  
  var ui = thisApp.getUi(); 
  var htmlOutput;  
  var setupCalled = "0";

  var lock = LockService.getScriptLock();
  var approvalState = PropertiesService.getDocumentProperties().getProperty('approvalState');
  var triggerState = PropertiesService.getDocumentProperties().getProperty('triggerState');
  setupCalled = PropertiesService.getDocumentProperties().getProperty('setupCalled');
  lock.releaseLock();

  if (setupCalled != "yes") {
    setupFunction("menuEdited"); 
  }
  
  if ( approvalState == 'Published') {
  thisApp.getUi() // Or DocumentApp or FormApp.
     .alert('You can\'t Approve a Published Doc!');
    return null;
  } else if ( approvalState == 'Needs Change') {
    thisApp.getUi() // Or DocumentApp or FormApp.
    .alert('You can\'t Approve a Doc that Needs Change without Author submitting for review!');
    return null;
  } else {

    if (approverEmail != null && approverEmail != undefined) {
      if (Session.getActiveUser().getEmail() == approverEmail) {
        findUpdateStatus(Session.getActiveUser().getEmail(), "Approved", 'reviewerListObj');
        sendEmails(Session.getActiveUser().getEmail(), 'Approved');
      } else {
        //findUpdateStatus(Session.getActiveUser().getEmail(), "Approved", 'reviewerListObj');
        
        thisApp.getUi() // Or DocumentApp or FormApp.
        .alert('You can\'t Approve for someone else!');
        return null;
      }
    } 
    
    Logger.log("From menuApprove: " + approverEmail);
    
    //if ((approverEmail == null || approverEmail == undefined) && (Session.getActiveUser().getEmail() ==  DriveApp.getFileById(ss.getId()).getOwner().getEmail())) { // If I own the doc I can change to any state
    if (Session.getActiveUser().getEmail() ==  DriveApp.getFileById(ss.getId()).getOwner().getEmail()) {  
      //findUpdateStatus(Session.getActiveUser().getEmail(), "Approved", 'reviewerListObj');
      
      sendEmails(Session.getActiveUser().getEmail(), 'Approved');
      
      lock = LockService.getScriptLock();
      PropertiesService.getDocumentProperties().setProperty('approvalState', 'Approved');
      lock.releaseLock();
    } else if ((Session.getActiveUser().getEmail() !=  DriveApp.getFileById(ss.getId()).getOwner().getEmail()) && (approverEmail == null || approverEmail === undefined)) { // This came in via the menu

      //PropertiesService.getDocumentProperties().setProperty('approvalState', 'Approved');
      sendEmails(Session.getActiveUser().getEmail(), 'Approved');

    }

    refreshSideBar();    
    if(triggerState == "unset") {
      setupTriggers(thisAppString, ss);      
    }
  }
  Logger.log("From menuApprove: Returning Approved");
  return "Approved";
}

/*
नित्येष्वभावादानित्येषु भवात्कारणे कालाख्येति ॥२।२।९॥ ~ Kaṇāda, Circa 500 B.C.E.
In the eternals non-existing and in the non-eternals existing is why time so called

Once the document has been published the doc owner can change its state to 'Approved' or 'Needs Change' only after taking it back to 'Being Edited'
*/

function menuPublish(e) {
 
  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;

  var ss;
  if (thisAppString == "SpreadsheetApp") {
    ss = thisApp.getActiveSpreadsheet();
  } else if (thisAppString == "DocumentApp") {
    ss = thisApp.getActiveDocument();
  } else if (thisAppString == "FormApp") {
    ss = thisApp.getActiveForm();
  } else if (thisAppString == "SlidesApp") {
    ss = thisApp.getActivePresentation();
  }
  var ui = thisApp.getUi();
 
  var htmlOutput;
  
  var setupCalled = "0";
  
  var lock = LockService.getScriptLock();
  var approvalState = PropertiesService.getDocumentProperties().getProperty('approvalState');
  var triggerState = PropertiesService.getDocumentProperties().getProperty('triggerState');
  setupCalled = PropertiesService.getDocumentProperties().getProperty('setupCalled');
  lock.releaseLock();

  if (setupCalled != "yes") {
    setupFunction("menuEdited"); 
  }
  
  if ( approvalState != 'Approved') {
    thisApp.getUi() // Or DocumentApp or FormApp.
    .alert('You can\'t Publish a Doc without Approval!');
  } else if (Session.getActiveUser().getEmail() !=  DriveApp.getFileById(ss.getId()).getOwner().getEmail()) {
    thisApp.getUi() // Or DocumentApp or FormApp.
        .alert('You can\'t Publish if you don\'t own the document!');
        return null;
  } else {
    
    lock = LockService.getScriptLock();
    PropertiesService.getDocumentProperties().setProperty('approvalState', 'Published');
    lock.releaseLock();

    refreshSideBar();
    sendEmails(Session.getActiveUser().getEmail(), 'Published', "menuPublish");
    
    if(triggerState == "unset") {
      setupTriggers(thisAppString, ss);
    }
  }
}
