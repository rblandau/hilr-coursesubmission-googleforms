# hilr-coursesubmission-googleforms

Contents: HILR GoogleForms javascript to email notifications to Curriculum Committee

## Abstract

The HILR Curriculum Committee (CC) collects and edits course proposals with Google Forms and Google Sheets.  The user proposing a course calls up a specific multi-page form, fills in the blanks and options, and submits the form.  The contents of the form are transferred to a single line in a Google Sheet.  The Committee may edit the proposal multiple times until it is either accepted for the next term or tabled.  

This JavaScript/GoogleScript software is attached to the form.  It sends an email to the CC's reflector, which forwards it to participants.  The need for this function is twofold.

1.  The CC members want to edit the proposal using the same form as used for submission.  Editing in a Google Sheet, particularly for long text items such as the course abstract, is extremely error-prone.

2.  The CC members need the editable URL for the form containing the latest contents of the submission, but we do not want the original submitter to be able to edit the contents after submitting the proposal the first time.  

The combination of these two requirements means that 

-   the Google Form cannot tell the submitter what the edit URL is for the proposal; and 

-   the CC editors need access to the edit URLs of all submissions.  

This software sends an email for every Submit event, whether the original submission or an edit.  That email contains the contents of the latest submission AND the edit URLs of all the submissions in the Google Sheet at this time.  

