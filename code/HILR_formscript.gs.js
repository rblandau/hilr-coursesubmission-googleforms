/**
 * HILR FormScript for Curriculum Committee submissions.
 * V04  20150309    RBLandau
 * V04x 20150814    RBLandau
 *                  Add a little debug info to old version, and update
 *                   the out-of-date version on the actual form.  
 * v05  20150814    RBLandau
 *                  New-ish version that almost works.  Still takes too much CPU time.  
 *                  Next: comment out all the old code that I think is no longer used 
 *                   and see if it still works.      
 * v06  20150815    Carefully log email construction and sending.  
 *                  In the future, we may have to worry about email length, which
 *                   Mr Google limits to 20KB.  Boy, I hope JavaScript Unicode 
 *                   doesn't count as two bytes per char.  Oops if so, already
 *                   in danger.  
 * v07  20150816    Remove most of the high-output logging to that it doesn't 
 *                   get truncated anymore.  
 *                  And don't include the log data at all; use a short stem value.  
 *                  Add dummy routines to test sending email.  
 *                  And comment out a lot of old code.
 * v08  20150817    Use MailApp instead of GmailApp today, works.  
 *                   Include little test functions for this.
 *                  Change From name on email to HILR something or other.
 *                  Finally remove the old code.  
 *                  Test some slight reformatting of the message.  
 * v09  20150929    Include the tail of the log file in the extrabody sent
 *                   to me.
 *                  Include partial course title as hint in the email 
 *                   subject line.
 */

// Yes, with no 'var' this is deliberately a global.  
// To be added to bottom of email.
sVersionNumber = '08';
sVersionDate = '20150817.1540'

/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/*
*************************************************************
STUFF THAT RICK ADDED: STOLEN, CRIBBED, AND MAYBE EVEN WORKS.
*************************************************************
*/


//-------------------------------------------------
// s e n d E m a i l A f t e r S u b m i t 
//-------------------------------------------------
function sendEmailAfterSubmit_RBL1(e) {
// Main entry point.  The Form Submit trigger should invoke this function.  

  //********************************************************************
  // Yes, with no 'var', this is a global.  Intentional in case needed elsewhere.  
//  bProductionVersion = false;       // <=========== EDIT THIS !!!!!! ===============
  bProductionVersion = true ;       // <=========== EDIT THIS !!!!!! ===============
  bIncludeLog = true;               // include the debug log in output messages.
  //********************************************************************

  if ( ! bProductionVersion ) {
    Logger.log("-------- TEST VERSION, NOT PRODUCTION ---------");
  } else {
    Logger.log("-------- PRODUCTION VERSION ---------");
  }

  Logger.log("ENTER sendEmailAfterSubmit e=%s", e);
  var oEventLists = fnoDumpEvent(e);
  var oItemDict = oEventLists.oItems;
  var lItemNames = oEventLists.lItemNames;
  
  //   U P P E R   B O D Y 
  
  var sUpperBody = "";
  //var sEditURL = returnResponseURL_RBL0(e);
  var sEditURL = oEventLists.sEditURL;
  sUpperBody += 'Editable URL with previous responses (below) filled in:\n' + sEditURL + '\n\n';
  var sTimestampBegin = getCurrentTimestamp_RBL1();
  sUpperBody += 'Begin: ' + sTimestampBegin + '\n\n';
  var maybeEditURL = sEditURL;

  for (var ii=0; ii<lItemNames.length; ii++) {
    var sQuestion = lItemNames[ii];
    var sAnswer = oItemDict[sQuestion];
    if (sAnswer != undefined)
    {
    	var sLine = '' + sQuestion + ' == "' + sAnswer + '"\n';
    	sUpperBody += sLine;
    }
  }

  sUpperBody += "\nEnd: " + getCurrentTimestamp_RBL1() + '\n';
  //Logger.log("UPPERBODY length|%s| completeUPPERBODY||%s||endUPPERBODY", sUpperBody.length, sUpperBody);

  //   L O W E R   B O D Y 

  var sLowerBody = "";
  var sURLList = "URLs for ALL responses, in row order.  Look at SGL and Title.\n\n";
  var asAllURLs = fnlGetAllUrls();
  sURLList += asAllURLs.join("\n") + "\n";
  sLowerBody += "\n" + sURLList; 
  var oForm = FormApp.getActiveForm();
  sLowerBody += '\nBeginning URL to blank form:\n' + oForm.getPublishedUrl() + '\n\n';
  sLowerBody += 'HILR CC Submision JS Script version ' + sVersionNumber + ' ' +  sVersionDate + '  \n';
  sLowerBody += '=======end=======\n';
  //Logger.log("LOWERBODY length|%s| completeLOWERBODY||%s||endLOWERBODY", sLowerBody.length, sLowerBody);

  //   E X T R A   B O D Y 

  // Special version of the body goes to the author; may include debug log.  
  // Optional because the log can be reeeaaallly looong, 1000s of lines.
  // Take only a largish tail of the log file.
  var sExtraBody = "";
  if (bIncludeLog) 
  {
    //var sLogData = Logger.getLog().toString();
    if (0) {
      var sLogData = "\n=\n==\n===\n   Short stand-in stem for log data, just to see what happens.\n===\n==\n=\n";
    } else {
      var sAllLogData = Logger.getLog().toString();
      var lLogData = sAllLogData.split("\n");
      var nLogLines = lLogData.length;
      // Get the tail of the log.  Today, fifty lines.
      var sLogData = lLogData.slice(nLogLines-50,nLogLines).join("\n");
    }
    sExtraBody += '\n\n====== Debug log data follows.  Please ignore from here to end. ======\n\n';
    sExtraBody += sLogData;
    sExtraBody += "\n ====== END of log data ======\n";
  }
  //Logger.log("EXTRABODY length|%s| completeEXTRABODY||%s||endEXTRABODY", sExtraBody.length, sExtraBody);

  //   F U L L   B O D Y 

  var sFullBody = sUpperBody + sLowerBody;
  var sLongBody = sFullBody + sExtraBody;
  //Logger.log("LONGBODY length|%s| completeLONGBODY||%s||endLONGBODY", sLongBody.length, sLongBody);

////////////// needs work ///////////////

  // Different email destinations depending on production vs test.
  if (! bProductionVersion)         // <=========== TEST VS PRODUCTION =============
  {
    // Rick's Howzit test version                       // T E S T 
    // Build new subject line.
    var oMess = oResp.oItems;     // object oMess[questionstring] = answerstring
    var sSglName = findItemResponseToKey("Howzit?",oMess);
    var sSglEMail = findItemResponseToKey("Why?",oMess);
    var nSubmNr = maybeSubmNumber; 
    var nRowNr = maybeRowNumber;
    var subject = "Row # {4} (submission {0}); from {1}; email={2}; at {3}".format( 
                   nSubmNr 
                 , sSglName
                 , sSglEMail
                 , sTimestampBegin
                 , nRowNr
               );
    Logger.log("EMAILSUBJECT complete|%s|",subject);
    // Get the email addresses of people who care.  
    var email_user = Session.getActiveUser().getEmail();
    var email2 = "landau@ricksoft.com";
    var email3 = "theatrewonk@gmail.com";
    var email4 = "receiver6_form@ricksoft.com";
    // For production, use only email2 and email3.  Comment out the others.  
    // Send an email with contents and a link to edit form with.
    //fnActuallySendEmail(email_user, subject, sFullBody);
    fnActuallySendEmail(email2, subject, sFullBody);
    fnActuallySendEmail(email3, subject, sFullBody);
    fnActuallySendEmail(email4, subject, sLongBody);
  } else
  {
    // HILR production version                          // P R O D U C T I O N
    // New subject line
    var oResp = oEventLists;
    var oMess = oResp.oItems;     // object oMess[questionstring] = answerstring
    // HILR production version
    var sSglName = findItemResponseToKey("SGL 1 Name",oMess);
    var sSglEMail = findItemResponseToKey("SGL 1 eMail",oMess);
    var sPartialTitle = findItemResponseToKey("Course Title",oMess).slice(0,24)+"...";
    var subject = "Submission from {0}; title={3}; email={1}; at {2}".format( 
                   sSglName
                 , sSglEMail
                 , sTimestampBegin
                 , sPartialTitle
               );
    Logger.log("EMAILSUBJECT |%s|",subject);
    // Get the email addresses of people who care.  
    var email_user = Session.getActiveUser().getEmail();
    var email2 = "hilr-cc-submissions@googlegroups.com";
    var email3 = "dickr@mac.com";
    var email4 = "receiver8_form@ricksoft.com";
    // For production, use only email2 and email3.  Comment out the others.  
    //  Oh, and email4 is a debug copy to me, best left turned on.  
    // Send an email with contents and a link to edit form with.
    //fnActuallySendEmail(email_user, subject, sFullBody);
    fnActuallySendEmail(email2, subject, sFullBody);    
    fnActuallySendEmail(email3, subject, sFullBody);
    fnActuallySendEmail(email4, subject, sLongBody);
  }

}

//-------------------------------------------------
// f n l G e t A l l U r l s 
//-------------------------------------------------
function fnlGetAllUrls() {
	/**
	Returns a string array of formatted URLs, in order from the sheet: 
	 asAllUrls
	*/
  Logger.log("ENTER fnlGetAllUrls ");

  var form = FormApp.getActiveForm();
  var formResponses = form.getResponses();
  var formResponsesLength = parseInt(formResponses.length);
  Logger.log("fnlGetAllUrls1 ROWN formResponses.length=%s", formResponsesLength );
  var asAllURLs = [];

  for (var i = 0; i < formResponsesLength; i++) {
    var formResponse = formResponses[i];
    var thisEditableURL = formResponse.getEditResponseUrl();
    //Logger.log("fnlGetAllURLs2 ROWN response i=%s, val=%s, url=%s", parseInt(i),formResponse,thisEditableURL);
    var dFormTimestamp = formResponse.getTimestamp();
    var sFormTimestamp = formatTimestamp_RBL1(dFormTimestamp);

		// Go thru all the items just to get the SGL Name and Course Title.
    var itemResponses = formResponse.getItemResponses();
    for (var j=0; j<itemResponses.length; j++) {
      var xitemResponse = itemResponses[j];
      var sQuestion = xitemResponse.getItem().getTitle();
      var sAnswer = xitemResponse.getResponse().toString().replace(/\s*$/,"");
      //Logger.log("fnlGetAllUrls3 ROWN resp i=%s j=%s itemresp=%s Q=%s A=%s", 
      //           parseInt(i),parseInt(j),xitemResponse,sQuestion,sAnswer );
      // Get the SGL name and course title for the long listing
      if ( ! bProductionVersion) 
      {                                     // T E S T 
        if ( sQuestion == "Howzit?" ) {
          sName = sAnswer;
        } else if ( sQuestion == "Why?" ) {
          sTitle = sAnswer;
        }
      } else
      {                                     // P R O D U C T I O N 
        if ( sQuestion == "SGL 1 Name" ) {
          sName = sAnswer;
        } else if ( sQuestion == "Course Title" ) {
          sTitle = sAnswer;
        }
      }
    }

    // Store all the info away in the arrays to be returned.  
    var nSub = i + 1;            // Humans are one-based, not zero-based.
    var nSubRow = nSub + 1;      // Table header in row 1, zero-th subm, called 1, is in row 2.
    //asURLs.push(thisEditableURL);
    //asNames.push(sName);
    //asTitles.push(sTitle);
    var sLineHead = "------ {1}: \"{2}\"".format(nSub,sName,sTitle,nSubRow,sFormTimestamp);
    var sLineURL = "{0}".format(thisEditableURL);
    var sLineTrailer = "Submission {0} row {3} at {4}.".format(nSub,sName,sTitle,nSubRow,sFormTimestamp);
    asAllURLs.push(sLineHead);
    asAllURLs.push(sLineURL);
    asAllURLs.push(sLineTrailer);
    asAllURLs.push(" ");
	}
	Logger.log("EXIT  fnlGetAllUrls result|%s|",asAllURLs);
	return(asAllURLs);
}

//-------------------------------------------------
// f o r m a t F o r m R e s p o n s e 
//-------------------------------------------------
function formatFormResponse_RBL4(myoResponse) {
/**
Returns a (dictionary) object containing
- asMessage: array of strings for the message, each string to be one line
- sResponse: row N+1 response N line for heading
- nResponse: integer response number
- sEmailUser: SGL1's declared email string
- oItemResponses: the itemResponses object from the formResponse object
- oItems: dictionary, oItems[question] = answer
*/
  Logger.log("ENTER formatFormResponse, arg|%s|",myoResponse);
  fnDumpObject(myoResponse,"formatFormResponse myoResponse arg in");
  // string format cheapo function from StackOverflow.
  if ( ! String.prototype.format ) {
    String.prototype.format = function() {
      var args = arguments;
      return this.replace(/{(\d+)}/g, function(match, number) { 
        return typeof args[number] != 'undefined'
        ? args[number]
        : match
        ;
      });
    }
  }
  
    var email_user = Session.getActiveUser().getEmail();
    var itemResponses = myoResponse.oItemResponses;
    var asMess = [];  // array of strings for the message, each a line
    var oItemlist = {};   // obj with properties obj[question]=answer.
    asMess.push('Contents of the most recent form submission, from ActiveUser = ' 
             + email_user + '\n');
    asMess.push('======'); 
    asMess.push( "submission at {0}:\n".format(myoResponse.sRespTimestamp) );
    var oReturnMe = fnoDeepCopy(myoResponse);
    oReturnMe.sResponse = "submitted at {0}:".format(myoResponse.sRespTimestamp);
    oReturnMe.oItemResponses = itemResponses;
    oReturnMe.sEmailUser = email_user;

    for (var jj = 0; jj < itemResponses.length; jj++ ) {
      var itemResponse = itemResponses[jj];
      var sAns = '{1} :: "{2}"';  // part of msg
      var sQuestion = itemResponse.getItem().getTitle();
      var sAnswer = itemResponse.getResponse()
      sAns = sAns.format( 
          (-887).toString(), 
          sQuestion, 
          sAnswer,
          (jj).toString(),
          (-888).toString()
          );
      asMess.push(sAns);
      oItemlist[sQuestion] = sAnswer;
    }  //endfor itemresponses
  asMess.push('======'); 
  oReturnMe.asMessage = asMess;
  oReturnMe.oItems = oItemlist;
  Logger.log("itemlist=|%s|", oReturnMe.oItems);

  fnDumpObject(oReturnMe,"formatFormResponse oReturnMe exit"); 
  Logger.log("EXIT  formatFormResponse, result|%s|",oReturnMe);
  return oReturnMe;
}

//-------------------------------------------------
// f n o D u m p E v e n t 
//-------------------------------------------------
/**
Returns a dictionary object from formatFormResponse().  
- asMessage: array of strings for the message, each string to be one line
- sResponse: row N+1 response N line for heading
- nResponse: integer response number
- sEmailUser: SGL1's declared email string
- oItemResponses: the itemResponses object from the formResponse object
- oItems: dictionary, oItems[question] = answer
plus
- sEmailUser: SGL1's declared email string
- sEditURL: editable URL for this event
- lItemNames: array (list) of question strings
*/
// For debugging, log the event contents, as much as we can see them.
function fnoDumpEvent(event1) {
  Logger.log("ENTER fnoDumpEvent, arg|%s|",event1);
  fnDumpObject(event1,"dumpEvent args ");
  var eResp = event1.response;
  var tmpEditURL = eResp.getEditResponseUrl();

  var oReturnMe = {};
  var lItemResponses = eResp.getItemResponses();
  event1.oItemResponses = lItemResponses;
  Logger.log("DUMPEV event1=%s number of itemResponses=%s editURL=%s", event1,lItemResponses.length,tmpEditURL);

  oReturnMe = formatFormResponse_RBL4(event1);
  oReturnMe.sEditURL = tmpEditURL;
  var oItemlist = {};
  var dRespTimestamp = eResp.getTimestamp();
  var sRespTimestamp = formatTimestamp_RBL1(dRespTimestamp);
  oItemlist[dRespTimestamp] = dRespTimestamp;
  oItemlist[sRespTimestamp] = sRespTimestamp;
  for (var sQuestion in eResp.oItems) 
  {
    var sAnswer = eResp.oItems[sQuestion];
    oItemlist[sQuestion] = sAnswer;
    Logger.log("formatted event Q|%s| A|%s|",sQuestion,sAnswer);
  }

  for (var j=0; j<lItemResponses.length; j++) {
    var xitemResponse = lItemResponses[j];
    var sQuestion = xitemResponse.getItem().getTitle();
    var sAnswer = xitemResponse.getResponse()
    oItemlist[sQuestion] = sAnswer;
    Logger.log("DUMPEV resp j|%s| Q|%s| A|%s|", 
               parseInt(j),sQuestion,sAnswer );
    // Dump details of the item (Class Item).
    xItem = xitemResponse.getItem();
    /*xItemText = xItem.asTextItem();*/
    xItemIndex = xItem.getIndex();
    xItemId = xItem.getId();
    xItemType = xItem.getType();
    xItemTitle = xItem.getTitle();
    //Logger.log("DUMPEV index|%s| id|%s| type|%s| title|%s|", xItemIndex,xItemId,xItemType,xItemTitle);
  }
  var lItemNames = [];
  for ( var q in oItemlist ) lItemNames.push(q);
  oReturnMe.oItemlist = oItemlist;
  oReturnMe.lItemNames = lItemNames;
  Logger.log("EXIT  fnoDumpEvent return|%s|",oReturnMe);
  return oReturnMe;
}

//-------------------------------------------------
// g e t C u r r e n t T i m e s t a m p 
//-------------------------------------------------
function getCurrentTimestamp_RBL1() {
  var now = new Date();
  var result = formatTimestamp_RBL1(now);
  return result;
}

//-------------------------------------------------
// f o r m a t T i m e s t a m p 
//-------------------------------------------------
function formatTimestamp_RBL1(when) {
  var date = [ when.getFullYear(), when.getMonth() + 1, when.getDate() ];
  // If months and days are less than 10, add a zero
  for ( var i = 1; i < 3; i++ ) {
      date[i] = ("00" + date[i]).substr(-2);
  }
  var sDate = date.join('-');

  var time = [ when.getHours(), when.getMinutes(), when.getSeconds() ];
  // Same for time parts.  
  for ( var i = 0; i < 3; i++ ) {
    time[i] = ("00" + time[i]).substr(-2);  // easier
  }
  sTime = time.join(":"); 
  var sMils = ("000" + when.getMilliseconds()).substr(-3);
  
  // Return the formatted string
  return sDate + '_' + sTime + "." + sMils;
}

//-------------------------------------------------
// f i n d I t e m R e s p o n s e T o K e y 
//-------------------------------------------------
function findItemResponseToKey(sKeyToFind,oItemdict) {
    var sValue = 'vvvvvvv';
    sValue = oItemdict[sKeyToFind];
    return sValue;    
}

//-------------------------------------------------
// f n o D e e p C o p y 
//-------------------------------------------------
function fnoDeepCopy(myoSomething) {
  var oNewthing = {};
  for ( var thing in myoSomething ) {
    oNewthing[thing] = myoSomething[thing];
  }
  return oNewthing;
}

//-------------------------------------------------
// f n o D u m p O b j e c t 
//-------------------------------------------------
function fnDumpObject(myoSomething,mysID) {
  for ( var thing in myoSomething ) {
    var thingval = myoSomething[thing];
    Logger.log("fnDumpObject ID|%s| key|%s|=val|%s|",mysID,thing,thingval);
  }
}



//-------------------------------------------------
// d u m m y F u n c t i o n ( )  
//-------------------------------------------------
function dummyFunction() {
// This function does nothing. 
// It is here so that I can run a dummy function to force authorization of the script.  
// Grumble, grumble, Mr Google.  

}

//-------------------------------------------------
// f n A c t u a l l y S e n d E m a i l 
//-------------------------------------------------
function fnActuallySendEmail(mysTarget, mysSubject, mysBody) {
  Logger.log("ENTER fnActuallySendEmail to|%s| subj|%s| bodylength|%s| beginBODY||%s||BODYend", mysTarget, mysSubject, mysBody.length, mysBody);
  var result = MailApp.sendEmail(mysTarget, mysSubject, mysBody,{name:"HILR CC Course Submission Form"});
  Logger.log("EXIT  fnActuallySendEmail to|%s| subj|%s| result|%s|", mysTarget, mysSubject, result);
}

//-------------------------------------------------
// f n T e s t E m a i l _ S m a l l 
//-------------------------------------------------
function fnTestEmail_Small() {
var to = "test7s@ricksoft.com";
var subj = "Test email from HILR form script - small";
var body = "This is the smallest test message.";
var result = fnActuallySendEmail(to,subj,body,{name:"fnTestEmail_Small"});
}

//-------------------------------------------------
// f n T e s t E m a i l _ M e d i u m 
//-------------------------------------------------
function fnTestEmail_Medium() {
var to = "test7m@ricksoft.com";
var subj = "Test email from HILR form script - medium";
var body = "This is the a less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
body += "This is the less small test message.\n";
var result = fnActuallySendEmail(to,subj,body,{name:"fnTestEmail_Medium"});
}


































/* 
        O L D   L E F T O V E R S 
Someday I will study these enough to use them as a framework,
 because they probably do some other useful things that I don't
 know about.
*/



/**
 * A global constant String holding the title of the add-on. This is
 * used to identify the add-on in the notification emails.
 */
var ADDON_TITLE = 'Form Notifications';

/**
 * A global constant 'notice' text to include with each email
 * notification.
 */
var NOTICE = "Form Notifications was created as an sample add-on, and is meant for \
demonstration purposes only. It should not be used for complex or important \
workflows. The number of notifications this add-on produces are limited by the \
owner's available email quota; it will not send email notifications if the \
owner's daily email quota has been exceeded. Collaborators using this add-on on \
the same form will be able to adjust the notification settings, but will not be \
able to disable the notification triggers set by other collaborators.";


 /**
  * Adds a custom menu to the active form to show the add-on sidebar.
  *
  * @param {object} e The event parameter for a simple onOpen trigger. To
  *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
  *     running in, inspect e.authMode.
  */
function onOpen(e) {
  FormApp.getUi()
      .createAddonMenu()
      .addItem('Configure notifications', 'showSidebar')
      .addItem('About', 'showAbout')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the form containing the add-on's user interface for
 * configuring the notifications this add-on will produce.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Form Notifications');
  FormApp.getUi().showSidebar(ui);
}

/**
 * Opens a purely-informational dialog in the form explaining details about
 * this add-on.
 */
function showAbout() {
  var ui = HtmlService.createHtmlOutputFromFile('About')
      .setWidth(420)
      .setHeight(270);
  FormApp.getUi().showModalDialog(ui, 'About Form Notifications');
}

/**
 * Save sidebar settings to this form's Properties, and update the onFormSubmit
 * trigger as needed.
 *
 * @param {Object} settings An Object containing key-value
 *      pairs to store.
 */
function saveSettings(settings) {
  PropertiesService.getDocumentProperties().setProperties(settings);
  adjustFormSubmitTrigger();
}

/**
 * Queries the User Properties and adds additional data required to populate
 * the sidebar UI elements.
 *
 * @return {Object} A collection of Property values and
 *     related data used to fill the configuration sidebar.
 */
function getSettings() {
  var settings = PropertiesService.getDocumentProperties().getProperties();

  // Use a default email if the creator email hasn't been provided yet.
  if (!settings.creatorEmail) {
    settings.creatorEmail = Session.getEffectiveUser().getEmail();
  }

  // Get text field items in the form and compile a list
  //   of their titles and IDs.
  var form = FormApp.getActiveForm();
  var textItems = form.getItems(FormApp.ItemType.TEXT);
  settings.textItems = [];
  for ( var i = 0; i < textItems.length; i++ ) {
    settings.textItems.push({
      title: textItems[i].getTitle(),
      id: textItems[i].getId()
    });
  }
  return settings;
}

/**
 * Adjust the onFormSubmit trigger based on user's requests.
 */
function adjustFormSubmitTrigger() {
  var form = FormApp.getActiveForm();
  var triggers = ScriptApp.getUserTriggers(form);
  var settings = PropertiesService.getDocumentProperties();
  var triggerNeeded =
      settings.getProperty('creatorNotify') == 'true' ||
      settings.getProperty('respondentNotify') == 'true';

  // Create a new trigger if required; delete existing trigger
  //   if it is not needed.
  var existingTrigger = null;
  for ( var i = 0; i < triggers.length; i++ ) {
    if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT) {
      existingTrigger = triggers[i];
      break;
    }
  }
  if (triggerNeeded && !existingTrigger) {
    var trigger = ScriptApp.newTrigger('respondToFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
  } else if (!triggerNeeded && existingTrigger) {
    ScriptApp.deleteTrigger(existingTrigger);
  }
}

/**
 * Responds to a form submission event if a onFormSubmit trigger has been
 * enabled.
 *
 * @param {Object} e The event parameter created by a form
 *      submission; see
 *      https://developers.google.com/apps-script/understanding_events
 */
function respondToFormSubmit(e) {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger require authorizations that have not
  // been supplied yet -- if so, warn the active user via email (if possible).
  // This check is required when using triggers with add-ons to maintain
  // functional triggers.
  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to reauthorize; the normal trigger action is not
    // conducted, since it authorization needs to be provided first. Send at
    // most one 'Authorization Required' email a day, to avoid spamming users
    // of the add-on.
    sendReauthorizationRequest();
  } else {
    // All required authorizations has been granted, so continue to respond to
    // the trigger event.

    // Check if the form creator needs to be notified; if so, construct and
    // send the notification.
    if (settings.getProperty('creatorNotify') == 'true') {
      sendCreatorNotification();
    }

    // Check if the form respondent needs to be notified; if so, construct and
    // send the notification. Be sure to respect the remaining email quota.
    if (settings.getProperty('respondentNotify') == 'true' &&
        MailApp.getRemainingDailyQuota() > 0) {
      sendRespondentNotification(e.response);
    }
  }
}


/**
 * Called when the user needs to reauthorize. Sends the user of the
 * add-on an email explaining the need to reauthorize and provides
 * a link for the user to do so. Capped to send at most one email
 * a day to prevent spamming the users of the add-on.
 */
function sendReauthorizationRequest() {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var lastAuthEmailDate = settings.getProperty('lastAuthEmailDate');
  var today = new Date().toDateString();
  if (lastAuthEmailDate != today) {
    if (MailApp.getRemainingDailyQuota() > 0) {
      var template =
          HtmlService.createTemplateFromFile('AuthorizationEmail');
      template.url = authInfo.getAuthorizationUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
          'Authorization Required',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    }
    settings.setProperty('lastAuthEmailDate', today);
  }
}

/**
 * Sends out creator notification email(s) if the current number
 * of form responses is an even multiple of the response step
 * setting.
 */
function sendCreatorNotification() {
  var form = FormApp.getActiveForm();
  var settings = PropertiesService.getDocumentProperties();
  var responseStep = settings.getProperty('responseStep');
  responseStep = responseStep ? parseInt(responseStep) : 10;

  // If the total number of form responses is an even multiple of the
  // response step setting, send a notification email(s) to the form
  // creator(s). For example, if the response step is 10, notifications
  // will be sent when there are 10, 20, 30, etc. total form responses
  // received.
  if (form.getResponses().length % responseStep == 0) {
    var addresses = settings.getProperty('creatorEmail').split(',');
    if (MailApp.getRemainingDailyQuota() > addresses.length) {
      var template =
          HtmlService.createTemplateFromFile('CreatorNotification');
      template.summary = form.getSummaryUrl();
      template.responses = form.getResponses().length;
      template.title = form.getTitle();
      template.responseStep = responseStep;
      template.formUrl = form.getEditUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      MailApp.sendEmail(settings.getProperty('creatorEmail'),
          form.getTitle() + ': Form submissions detected',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    }
  }
}

/**
 * Sends out respondent notificiation emails.
 *
 * @param {FormResponse} response FormResponse object of the event
 *      that triggered this notification
 */
function sendRespondentNotification(response, aboutText) {
  var form = FormApp.getActiveForm();
  var settings = PropertiesService.getDocumentProperties();
  var emailId = settings.getProperty('respondentEmailItemId');
  var emailItem = form.getItemById(parseInt(emailId));
  var respondentEmail = response.getResponseForItem(emailItem)
      .getResponse();
  if (respondentEmail) {
    var template =
        HtmlService.createTemplateFromFile('RespondentNotification');
    template.paragraphs = settings.getProperty('responseText').split('\n');
    template.notice = NOTICE;
    var message = template.evaluate();
    MailApp.sendEmail(respondentEmail,
        'Thank you for filling out form ' + form.getTitle() + '!',
        message.getContent(), {
          name: form.getTitle(),
            htmlBody: message.getContent()
        });
  }
}



//END
