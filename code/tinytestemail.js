/*
Tiny test to see if we can ever send any email from a form script.
*/

//-------------------------------------------------
// d u m m y F u n c t i o n  
//-------------------------------------------------
function dummyFunction() {
// This function does nothing. 
// It is here so that I can run a dummy function to force authorization of the script.  
}

//-------------------------------------------------
// f n A c t u a l l y S e n d E m a i l 
//-------------------------------------------------
function fnActuallySendEmail(mysTarget, mysSubject, mysBody) {
  Logger.log("ENTER fnActuallySendEmail to|%s| subj|%s| beginBODY||%s||BODYend", mysTarget, mysSubject, mysBody);
  var result = MailApp.sendEmail(mysTarget, mysSubject, mysBody,{name:"HILR CC Course Submission Form"});
  Logger.log("EXIT  fnActuallySendEmail to|%s| subj|%s| body||%s|| result|%s|", mysTarget, mysSubject, mysBody, result);
}

//-------------------------------------------------
// f n T e s t E m a i l _ S m a l l 
//-------------------------------------------------
function fnTestEmail_Small() {
var to = "test7s@ricksoft.com";
var subj = "Test email from tinytest script - small";
var body = "This is a small test message.";
var result = fnActuallySendEmail(to,subj,body);
}

/* Execution transcript: 
[15-08-17 00:22:31:444 EDT] Starting execution
[15-08-17 00:22:31:450 EDT] Logger.log([ENTER fnActuallySendEmail to|%s| subj|%s| beginBODY||%s||BODYend, [test7s@ricksoft.com, Test email from tinytest script - small, This is a small test message.]]) [0 seconds]
[15-08-17 00:22:31:466 EDT] GmailApp.sendEmail([test7s@ricksoft.com, Test email from tinytest script - small, This is a small test message.]) [0.015 seconds]
[15-08-17 00:22:31:596 EDT] Execution failed: Authorization is required to perform that action. [0.019 seconds total runtime]
*/
/* Log
[15-08-17 00:22:31:449 EDT] ENTER fnActuallySendEmail to|test7s@ricksoft.com| subj|Test email from tinytest script - small| beginBODY||This is a small test message.||BODYend
/*
