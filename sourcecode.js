//A generic function to read a value from a particular cell of a given sheet.
//Parameter(s) - Sheet - Reference to the sheet.
//               row - absolute # of the row
//               column - absolute # of column
function getValueFromCell(sheet, row, column) {
  var cellValue = sheet.getRange(row,column).getValue()
  return cellValue
}

//A generic function to get the attachment id of a given link.
//User needs to copy the link from drive and paste it in given cell
function getAttachmentId(link){
  
  var replacedStr = link.replace("https://drive.google.com/open?id=","");
  return replacedStr;
}

//A function to populate the variables required for sending mail
function getGlobals(){
  bodyTemplate = "Hi {{First_Name}}!<br><p>I am pleased to invite you to our meeting number # {{Meet_Number}}.<br>"
  + "Please find attached the agenda for the upcoming meeting.<br>"
  + "Should you require further clarification, please do not hesitate to contact me.<br><br>"
  + "Thanks and Regards<br>Tushar Saurabh,<br>VP-Membership,<br>Medley Toastmasters Club.<br>Phone: 805-602-4308<br>"
  + "Like us on FB: http://www.facebook.com/MedleyToastmasters"
  
  spreadSheet = SpreadsheetApp.openById("1Xde65pEGL0KlGlnglE9LGjjSVGV15cuZAOVAosk71ak");
  mailDetails = spreadSheet.getSheetByName("Mail Merge");
  mailTrigger = spreadSheet.getSheetByName("Meeting Info");
  
  meetingNumber = getValueFromCell(mailTrigger,11,6);
  
  
  meetingPlace =  getValueFromCell(mailTrigger,12,6);
  attachmentLink = getValueFromCell(mailTrigger,13,6);
  
  attachmentID = getAttachmentId(attachmentLink);
  
  data = mailDetails.getDataRange().getValues();
  
  file = DriveApp.getFileById(attachmentID);
  
  emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  
  Logger.log("Remaining Quota :" + emailQuotaRemaining);
}

function entryPoint(){
  
  getGlobals();
  
  var previousValue = "";
  var mailCount = 0;
  
  for (var i = data.length-1; i > 0; i--) {
    
    bodyWithFirstName = bodyTemplate.replace("{{First_Name}}",data[i][0]);
    mailBody = bodyWithFirstName.replace("{{Meet_Number}}",meetingNumber);
    
    //The crux of the app. Currently, a user can send only 100 mails through App Script. Hence after 100 limit it reached, the script will just
    //a trigger that will be created for each 100 entries read. The trigger will be seprated by 24 Hrs.
    
    //I have added 100 to i, since the mail is sent from last entry and hence it was done to fix the bug, that caused only one trigger to set.
    
    if (mailCount > emailQuotaRemaining){
      var sendMailAfter = 24 * Math.max(Math.floor(((i+100) /100)),1)*60 * 60 * 1000
      if (sendMailAfter != previousValue){
        
        ScriptApp.newTrigger("sendScheduledMail")
        .timeBased()
        .after(sendMailAfter)
        .create();
      }
      previousValue = sendMailAfter;
      mailDetails.getRange(1 + i, 7).setValue("EMAIL_SCHEDULED");
    } else {
    
//      if (i == 1){  
      if (data[i][2] != ""){
        mailCount = mailCount + 1;
        try{
          MailApp.sendEmail({
            //          to: "XXXX@gmail.com",
            to:data[i][2],
            subject: "Welcome to Medley Meeting",
            htmlBody: mailBody,
            attachments: file.getAs(MimeType.PDF)
          });
          mailDetails.getRange(1 + i, 7).setValue("EMAIL_SENT");
        } catch (Err){
          Logger.log(Err.message);
          mailDetails.getRange(1 + i, 7).setValue("EMAIL_SENT_ERROR");
        }
      }
    }
  }
}

//The trigger calls the below function to send another 100 mail.
function sendScheduledMail(){
  
  getGlobals();
  
  var mailCount = 0;
  Logger.log("mail Detail " + mailDetails.getName());
  Logger.log("data Length " + data.length);
  for (var i = data.length - 1; i > 0; i--) {
    
    Logger.log("Email Status " + mailDetails.getRange(1 + i, 7).getValue());
    if (data[i][6] == "EMAIL_SCHEDULED" && data[i][2] != "" ){
      
      bodyWithFirstName = bodyTemplate.replace("{{First_Name}}",data[i][0]);
      mailBody = bodyWithFirstName.replace("{{Meet_Number}}",meetingNumber);
      
      if (emailQuotaRemaining > 0){
        try{
        MailApp.sendEmail({
     //     to: "tusharacc@gmail.com",
          to:data[i][2],
          subject: "Welcome to Medley Meeting",
          htmlBody: mailBody,
          attachments: file.getAs(MimeType.PDF)
        });
        mailCount = mailCount + 1;
        mailDetails.getRange(1 + i, 7).setValue("EMAIL_SENT");
        } catch(Err){
          Logger.log(Err.message);
          mailDetails.getRange(1 + i, 7).setValue("EMAIL_SENT_ERROR");
        }
      }
    }
    emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    Logger.log("Mail Sent:" + mailCount)
  }
}
        
function getRemainingQuota(){
  emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  
  Logger.log("Remaining Quota :" + emailQuotaRemaining);
  
}
 
