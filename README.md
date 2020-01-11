# Spreadsheet2gmail1

```
---sendEmails.gs---
function sendEmails() {
  
  // get spreadsheet data
  var ss = SpreadsheetApp.getActive();
  var dataSheet = ss.getSheetByName('Data');
  var dataValues = dataSheet.getDataRange().getValues();
  var dataLastRow = dataSheet.getLastRow();
  var timeZone = ss.getSpreadsheetTimeZone();
  
  ss.toast('Starting to send emails');
  
  // run Function to get email body from emailMessage sheet
  var emailContent = getEmailBody();
  var plainBody = emailContent['plainBody'];
  
  
  // loop through each row *************************************************
  for (var i=1; i<dataLastRow; i++) {
    
    var memberEmailAddress = dataValues[i][3];
    Logger.log('Member email address is: ' + memberEmailAddress);
    var emailSent = dataValues[i][4];
  
    // check email address not blank and no value in 'Email Sent' Column
    if ((memberEmailAddress != '') && (emailSent == '')) {
      
      var memberID = dataValues[i][0];
      var memberFullName = dataValues[i][1];
      var memberSurname = dataValues[i][2];

      
      // *********************** START OF EDIT EMAIL SECTION ***********************
      var subject = 'Bonsoir ' + memberSurname;    // 'Bonsoir ' + memberFullName[0] + ' ' + memberSurname;
      

      //replace <<tags>> with appropriate content
      
      var tempPlainBody1 = plainBody.replace('<<memberID>>', memberID);
      var tempPlainBody2 = tempPlainBody1.replace('<<FullName>>', memberFullName);
      var tempPlainBody3 = tempPlainBody2.replace('<<Surname>>', memberSurname);
      var newPlainBody = tempPlainBody3.replace(/\<br\/>/mg, ''); // replace all <br/> tags with nothing
     

      // edit any options for email below
      var options = {replyTo:'testemail@testmail.com'};
      
      // try/catch to prevent script error if invalid email address
      try {
        // send email
        MailApp.sendEmail(memberEmailAddress, subject, newPlainBody, options);
        // flag to confirm timestamp can be written for successful send
        var sendSuccess = true;
      }
      catch(e) {
        Logger.log('Error with email: ' + e);
        var sendSuccess = false;
      }
      // *********************** END OF EDIT EMAIL SECTION ***********************
      
      
      // write timestamp to confirm email been sent if Flag is true
      if (sendSuccess) {
      var date = new Date;
      var timestamp = Utilities.formatDate(date, timeZone, "dd/MM/yy @ HH:mm:ss");
      dataSheet.getRange(i+1, 5).setValue(timestamp);
      }
      
    }// end of check email address not blank, no value in 'Email Sent' Column
    else {Logger.log('Email was not sent for row: ' + (i+1))}
    
  } // end of loop through each row ******************************************

  
  ss.toast('Sending emails complete');
  
  
}


function getEmailBody() {
 
  // get spreadsheet data
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var emailMessageSheet = ss.getSheetByName('emailMessage');
  var bodyGreeting = emailMessageSheet.getRange(1, 2).getValue();
  var bodyStart = emailMessageSheet.getRange(2, 2).getValue();
  var bodyMiddle = emailMessageSheet.getRange(3, 2).getValue();
  var bodyEnd = emailMessageSheet.getRange(4, 2).getValue();
  var signOffStart = emailMessageSheet.getRange(5, 2).getValue();
  var signOffEnd = emailMessageSheet.getRange(6, 2).getValue();
  var footer = emailMessageSheet.getRange(7, 2).getValue();
  
  
  // create Plain body
  var plainBody = bodyGreeting + '\n \n';
  plainBody+= bodyStart + '\n \n';
  plainBody+= bodyMiddle + '\n \n';
  plainBody+= bodyEnd + '\n \n';
  plainBody+= signOffStart + '\n \n';
  plainBody+= signOffEnd + '\n \n';
  plainBody+= footer + '\n \n';
  
  // collate data to return
  var toReturn = {'plainBody':plainBody};
  
  return toReturn;
  
}


/* To add in the plain body

 plainBody+= 'Member ID: ' + '<<memberID>>' + '\n';
 plainBody+= 'Full Name: ' + '<<FullName>>' + '\n';
 plainBody+= 'Surname: ' + '<<Surname>>' + '\n';

*/
```
