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
  
  mailDetails = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mail Merge");
  mailTrigger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meeting Info");
  
  meetingNumber = getValueFromCell(mailTrigger,11,6);
  
  
  meetingPlace =  getValueFromCell(mailTrigger,12,6);
  attachmentLink = getValueFromCell(mailTrigger,13,6);
  
  attachmentID = getAttachmentId(attachmentLink);
  
  data = mailDetails.getDataRange().getValues();
  
  file = DriveApp.getFileById(attachmentID);
}

function entryPoint(){
  
  getGlobals();
  
  var previousValue = "";
  var mailCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    
    bodyWithFirstName = bodyTemplate.replace("{{First_Name}}",data[i][0]);
    mailBody = bodyWithFirstName.replace("{{Meet_Number}}",meetingNumber);
    
    //The crux of the app. Currently, a user can send only 100 mails through App Script. Hence after 100 limit it reached, the script will just
    //a trigger that will be created for each 100 entries read. The trigger will be seprated by 24 Hrs.
    
    if (mailCount > 100){
      var sendMailAfter = 24 * Math.floor((i /100))*60 * 60 * 100
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
        MailApp.sendEmail({
//          to: "XXXX@gmail.com",
            to:data[i][2],
          subject: "Welcome to Medley Meeting",
          htmlBody: mailBody,
          attachments: file.getAs(MimeType.PDF)
        });
        mailDetails.getRange(1 + i, 7).setValue("EMAIL_SENT");
        
      }
    }
  }
}

//The trigger calls the below function to send another 100 mail.
function sendScheduledMail(){
  
  getGlobals();
  
  var mailCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    
    if (mailDetails.getRange(1 + i, 7).getValue == "EMAIL_SCHEDULED"){
      
      bodyWithFirstName = bodyTemplate.replace("{{First_Name}}",data[i][0]);
      mailBody = bodyWithFirstName.replace("{{Meet_Number}}",meetingNumber);
      
      if (mailCount <= 100){
        MailApp.sendEmail({
          to: "tusharacc@gmail.com",
     //     to:data[i][2],
          subject: "Welcome to Medley Meeting",
          htmlBody: mailBody,
          attachments: file.getAs(MimeType.PDF)
        });
        mailCount = mailCount + 1;
        mailDetails.getRange(1 + i, 7).setValue("EMAIL_SENT");
      }
    }
  }
}
          
 
