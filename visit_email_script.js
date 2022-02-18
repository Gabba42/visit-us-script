function formSubmit(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = ss.getSheetByName('Form Responses 1'); 
    var timestamp = e.namedValues['Timestamp'][0];
    var respondentEmailAddress = e.namedValues['Email Address'][0]; 
    var respondentName = e.namedValues['Name'][0]; 
    var respondentPhoneNumber = e.namedValues['Phone Number'][0];
    var dateTimeSelection = e.namedValues['Date & Time Selection'][0];
    var reasonForVisit = e.namedValues['Reason for Visit'][0]; 
    
    //Tour variables 
    var tourInterest = e.namedValues['Tour Interest'][0];
    Logger.log("tourInterest: " + tourInterest); 
    
    //New Member Orientation variables
    var nmoInterest = e.namedValues['Membership Interest'][0];
    var moneyAck = e.namedValues['Access Card Deposit Agreement'][0]; 
    Logger.log("nmoInterest: " + nmoInterest);
    
    //Other variables
    var otherInterest = e.namedValues['Other Reason'][0]; 
    Logger.log("otherInterest: " + otherInterest);
    
    //email template for Admin notification
    var toAdminEmailAdress = 'contact@pikespeakmakerspace.org';  
    var emailSubject = respondentName + '_' + reasonForVisit + '_Visit-Us'; 
    var textBody =
        'There is a new submission for your Visit Us Form!\n' +
        'Form Details:\n';
    
    if(reasonForVisit === 'Take A Tour') {
        textBody += 
            'Date/Time Submitted: ' + timestamp + '\n' +
            'Submitted by: ' + respondentName + '\n' + 
            'Email: ' + respondentEmailAddress + '\n' +
            'Phone Number: ' + respondentPhoneNumber + '\n' +   
            'Reason for visit: ' + reasonForVisit + '\n' + 
            'Date & Time: ' + dateTimeSelection + '\n' + 
            'Their interest in PPM: ' + tourInterest + '\n';
      } else if (reasonForVisit === 'New Member Safety Orientation') {
        textBody += 
            'Date/Time Submitted: ' + timestamp + '\n' +
            'Submitted by: ' + respondentName + '\n' + 
            'Email: ' + respondentEmailAddress + '\n' +
            'Phone Number: ' + respondentPhoneNumber + '\n' +   
            'Reason for visit: ' + reasonForVisit + '\n' + 
            'Date & Time: ' + dateTimeSelection + '\n' + 
            'Their interest in PPM: ' + nmoInterest + '\n' + 
            'Money Ack Response: ' + moneyAck + '\n';
      } else {
        textBody += 
            'Date/Time Submitted: ' + timestamp + '\n' +
            'Submitted by: ' + respondentName + '\n' + 
            'Email: ' + respondentEmailAddress + '\n' +
            'Phone Number: ' + respondentPhoneNumber + '\n' +   
            'Reason for visit: ' + reasonForVisit + '\n' + 
            'Date & Time: ' + dateTimeSelection + '\n' + 
            'Their interest in PPM: ' + otherInterest + '\n'; 
      }
    
    var htmlText = 
        '<h3>There is a new submission for your Visit Us Form!</h3>' + 
        '<h4>Form Details:</h4>'; 
    
      if(reasonForVisit === 'Take A Tour') {
        htmlText += 
          '<ul>' + 
            '<li>Date/Time Submitted: ' + timestamp + '</li>' +
            '<li>Submitted by: ' + respondentName + '</li>' + 
            '<li>Email: ' + respondentEmailAddress + '</li>' +
            '<li>Phone Number: ' + respondentPhoneNumber + '</li>' + 
            '<li>Reason for visit: ' + reasonForVisit + '</li>' + 
            '<li>Date & Time: ' + dateTimeSelection + '</li>' + 
            '<li>Their interest in PPM: ' + tourInterest + '</li>'
        '</ul>';
      } else if (reasonForVisit === 'New Member Safety Orientation') {
        htmlText += 
          '<ul>' + 
            '<li>Date/Time Submitted: ' + timestamp + '</li>' +
            '<li>Submitted by: ' + respondentName + '</li>' + 
            '<li>Email: ' + respondentEmailAddress + '</li>' +
            '<li>Phone Number: ' + respondentPhoneNumber + '</li>' + 
            '<li>Reason for visit: ' + reasonForVisit + '</li>' + 
            '<li>Date & Time: ' + dateTimeSelection + '</li>' + 
            '<li>Their interest in PPM: ' + nmoInterest + '</li>' + 
            '<li>Money Ack Response: ' + moneyAck + '</li>'
        '</ul>';
      } else {
        htmlText += 
          '<ul>' + 
            '<li>Date/Time Submitted: ' + timestamp + '</li>' +
            '<li>Submitted by: ' + respondentName + '</li>' + 
            '<li>Email: ' + respondentEmailAddress + '</li>' +
            '<li>Phone Number: ' + respondentPhoneNumber + '</li>' + 
            '<li>Reason for visit: ' + reasonForVisit + '</li>' + 
            '<li>Date & Time: ' + dateTimeSelection + '</li>' + 
            '<li>Their interest in PPM: ' + otherInterest + '</li>'
        '</ul>';
      }
    
    var options = {
      htmlBody: htmlText
    }
    
    //send email notification to Admin
    GmailApp.sendEmail(toAdminEmailAdress, emailSubject, textBody, options);
    
    
    //Email template to Respondent
    var confirmationSubject = 'Confirmation of Your Visit to Pikes Peak Makerspace!'; 
    var confirmationBody = 'The email requires HTML support. Please use a client that supports HTML';
    var confirmationHtmlText = 
        '<h3>Below are the details of your visit with us!</h3>' + 
        '<h4>Visit Details:</h4>'; 
    
     if(reasonForVisit === 'Take A Tour') {
        confirmationHtmlText += 
          '<ul>' + 
            '<li>Date/Time Submitted: ' + timestamp + '</li>' +
            '<li>Submitted by: ' + respondentName + '</li>' + 
            '<li>Email: ' + respondentEmailAddress + '</li>' +
            '<li>Reason for visit: ' + reasonForVisit + '</li>' + 
            '<li>Date & Time: ' + dateTimeSelection + '</li>' + 
            '<li>Your interest in PPM: ' + tourInterest + '</li>'
          '</ul>' + 
         '<br>';
      } else if (reasonForVisit === 'New Member Safety Orientation') {
        confirmationHtmlText += 
          '<ul>' + 
            '<li>Date/Time Submitted: ' + timestamp + '</li>' +
            '<li>Submitted by: ' + respondentName + '</li>' + 
            '<li>Email: ' + respondentEmailAddress + '</li>' +
            '<li>Reason for visit: ' + reasonForVisit + '</li>' + 
            '<li>Date & Time: ' + dateTimeSelection + '</li>' + 
            '<li>Your interest in PPM: ' + nmoInterest + '</li>' + 
            '<li>Money Ack Response: ' + moneyAck + '</li>'
          '</ul>' + 
         '<br>';
      } else {
        confirmationHtmlText += 
          '<ul>' + 
            '<li>Date/Time Submitted: ' + timestamp + '</li>' +
            '<li>Submitted by: ' + respondentName + '</li>' + 
            '<li>Email: ' + respondentEmailAddress + '</li>' +
            '<li>Reason for visit: ' + reasonForVisit + '</li>' + 
            '<li>Date & Time: ' + dateTimeSelection + '</li>' + 
            '<li>Your interest in PPM: ' + otherInterest + '</li>'
           '</ul>' + 
         '<br>';
      }
        confirmationHtmlText +=
         '<br>' +
         '<h4>Need to Reschedule or Cancel?</h4>' + 
         '<p>If you need to change the date or time you have selected or you need to cancel your visit for any reason, you can reply off this email or call us directly.</p>' + 
         '<p>We work mainly on volunteers so no shows really hurt us! We kindly ask if you need to change or cancel your visit to please get in contact with us!</p>' +
         '<br>' + 
         '<h4>Please Note!</h4>' +
         '<p>Just as a reminder, there might be other community members joining you on your visit to us. Also, per the Colorado state-wide mandate, we kindly ask that you <strong>wear a mask</strong> during the entire duration of your visit.</p>' + 
         '<br>' +
         '<p>We look forward to meeting you!</p>' + 
         '<br>' + 
         '<p><strong>Pikes Peak Makerspace</strong></p>' +
         '<p>735 E Pikes Peak Ave</p>' +
         '<p>Colorado Springs, CO 80903</p>' +
         '<a href="tel:7194456253">719-445-MAKE (6253)</a>'; 
    
    var options = {
      htmlBody: confirmationHtmlText,
      name: 'Pikes Peak Makerspace', //cannot be Contact Us. Gmail gets confused and sends over the respondent's confirmation as well to us
      from: 'contact@pikespeakmakerspace.org',
      replyTo: 'contact@pikespeakmakerspace.org'
    }
    
    
    //send email confirmation to Respondent
    GmailApp.sendEmail(respondentEmailAddress, confirmationSubject, confirmationBody, options);
  }