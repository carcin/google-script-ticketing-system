//FUNCTION - ADD THE TICKET FROM THE FORM AND SEND EMAILS TO THE CUSTOMER AND SUPPORT
function onFormSubmit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responsesSheet = ss.getSheetByName("Odpovědi formuláře");
  
  //format the sheet with answers
  responsesSheet.getDataRange().setHorizontalAlignment("center")
                              .setWrap(true)
                              .setBorder(true, true, true, true, true, true)
                              .setVerticalAlignment("center");
  
  //getting last row and column
  var lastRow = responsesSheet.getLastRow();
  var lastColumn = responsesSheet.getLastColumn();
  
  //set Open status to the last ticket from the form
  var setStatus = responsesSheet.getRange(lastRow, 7).setValue("Open");
  
  //getting variables from the Answers
  var lastRowValues = responsesSheet.getRange(lastRow, 1, 7, lastColumn).getValues();
  var timeStamp = lastRowValues[0][0];
  var type = lastRowValues[0][1];
  var description = lastRowValues[0][2];
  var name = lastRowValues[0][3];
  var email = lastRowValues[0][4];
 
  //change the format of the date
  var timeZone = Session.getScriptTimeZone();
  var shortTimeStamp = Utilities.formatDate(timeStamp, timeZone, 'dd.MM.YYYY HH:mm');
  
  //URL of the Spreadsheet
  var ssURL = ss.getUrl();
  
  //email template with the ticket - for the support
  var emailSubject = "Požadavek na podporu AppSatori - z "+ shortTimeStamp;
  var emailBody = "<h3><u>Ticket</u></h3> \
<strong>Typ: </strong>"+ type +"<br /> \
<strong>Požadavek: </strong>"+ description +"<br /> \
</p><p><hr /> \
<strong>Od: </strong>"+ name +"<br /> \
<strong>Email: </strong>"+ email +"<br /> \
</p><p><hr /> \
<strong>Odkaz na log: </strong>"+ ssURL +"<br /> \
";
  
  //email template with the ticket - for the customer
  var copySubject = "Váš Požadavek na podporu AppSatori - z "+ shortTimeStamp;
  var copyBody = "<h3><u>Ticket</u></h3> \
Toto je kopie Vašeho požadavku na podporu AppSatori. Potvrzujeme, že jsme jej obdrželi a do 24 hodin se Vám ozveme.<br /> \
<strong>Typ: </strong>"+ type +"<br /> \
<strong>Požadavek: </strong>"+ description +"<br /> \
</p><p><hr /> \
<strong>Od: </strong>"+ name +"<br /> \
<strong>Email: </strong>"+ email +"<br /> \
</p><p><hr /> \
<p>V případě, že jste tento požadavek nezasílali, kontaktujte nás prosím okamžitě na 702 168 190 či 602 687 061.</p><br /> \
";

  //getting email adresses with support emails (in case of different support email addresses)
  var supportContacts = ss.getSheetByName("Emaily");
  var numEmailRows = supportContacts.getLastRow();
  
  //condition for different support email adresses
  if(type==="Licence (přidání dalších licencí, prodloužení, odebrání...)") {
    var emailTo = supportContacts.getRange(2, 2, numEmailRows, 1).getValues();
  } else {
    var emailTo = supportContacts.getRange(2, 1, numEmailRows, 1).getValues();
  }
  
  //sending emails to the support
  GmailApp.sendEmail(emailTo,emailSubject,emailBody, {htmlBody: emailBody, replyTo: "pomoc@appsatori.eu"});
  GmailApp.sendEmail(email,copySubject, copyBody, {htmlBody: copyBody, replyTo: "pomoc@appsatori.eu"})
}
  
//FUNCTION - ADD TODAY`S DATE IN CASE OF EDIT TICKET`S STATE
function onEdit(event) { 
  var timezone = "GMT+2";
  var timestamp_format = "dd-MM-YYYY";
  var updateColName = "Stav řešení";
  var timeStampColName = "Datum posledního updatu";
  var sheet = event.source.getSheetByName('Odpovědi formuláře');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responsesSheet = ss.getSheetByName("Odpovědi formuláře");
  var lastRow = responsesSheet.getLastRow();
  var lastColumn = responsesSheet.getLastColumn();

  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  
  //condition to edit the field with "last update date" and "notification"
  if (dateCol > -1 && index > 1 && editColumn == updateCol) {
    var lastUpdate = sheet.getRange(index, dateCol + 1); //Initialization of an active row and column "Last update date"
    var notification = sheet.getRange(index, dateCol + 2); //Initialization of an active row and column "Notification"
    var date = Utilities.formatDate(new Date(), "GMT+2", "dd.MM.YYYY HH:mm")
    lastUpdate.setValue(date); //set today`s date in column "Last update date"
    notification.setValue("") //erase "NO" in column "Notification"
  }
}

//FUNCTION - COMPARE TODAY`S DATE AND DATE OF THE LAST UPDATE
function notification() {
  var responsesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Odpovědi formuláře");
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  var lastRow = responsesSheet.getLastRow();
  var columnH = responsesSheet.getRange(1, 8, responsesSheet.getLastRow(), 1).getValues();
  var states = responsesSheet.getRange(1, 7, responsesSheet.getLastRow(), 1).getValues();
  var notification = responsesSheet.getRange(1, 9, responsesSheet.getLastRow(), 1).getValues();
  var today = new Date(new Date());
  Logger.log(today)
  
  //loop for comparing today`s date and date of the last update
  for (var i=1; i < columnH.length; i++) {
    var dataday =columnH[i][0];
    Logger.log(dataday)
    var specificState =states[i][0];
    var notificationState = notification[i][0];
    var notificationCol = responsesSheet.getRange(i+1, 9);//get column "Notification"
    
    //1. condition - last update is older than today and state is "Solving" and Notification doesn`t contain "NO"
    if (dataday < today && specificState == "Řeší se" && notificationState !== "NE"){
      {
        var lastColumn = responsesSheet.getLastColumn();
        var lastRowValues = responsesSheet.getRange(i+1, 1, 7, lastColumn).getValues();
        Logger.log(lastRowValues)
        var timeZone = "GMT+2";
        var timeStamp = lastRowValues[0][0];
        var shortTimeStamp = Utilities.formatDate(timeStamp, timeZone, "dd.MM.YYYY HH:mm");
        var type = lastRowValues[0][1];
        var description = lastRowValues[0][2];
        var name = lastRowValues[0][3];
        var email = lastRowValues[0][4];
        var note = lastRowValues[0][5];
        var emailSubject = "Požadavek na podporu AppSatori - z "+ shortTimeStamp + " JE V ŘEŠENÍ";
        var emailBody = "<i>Právě řešíme Váš požadavek s popisem: <br /> \
        <p>" + description + "</p></i> \
        </p><p><hr /> \
        <b>Popis řešení:</b> " + note + " \
        <p><i>S pozdravem</p> \
        Podpora AppSatori <br />\
        pomoc@appsatori.eu \
        ";
        
        //send email to the customer
        GmailApp.sendEmail(email,emailSubject,emailBody, {htmlBody: emailBody, replyTo: "pomoc@appsatori.eu"});
        
        //sets value "NO" to the "Notification" column
        notificationCol.setValue("NE").setVerticalAlignment("center");
      }
    }
    
    //2. condition - last update is older than today and state is "Waiting for the customer" and Notification doesn`t contain "NO"
    else if (dataday < today && specificState == "Čeká se na reakci zákazníka" && notificationState !== "NE"){
      {
        var lastColumn = responsesSheet.getLastColumn();
        var lastRowValues = responsesSheet.getRange(i+1, 1, 7, lastColumn).getValues();
        Logger.log(lastRowValues)
        var timeZone = "GMT+2";
        var timeStamp = lastRowValues[0][0];
        var shortTimeStamp = Utilities.formatDate(timeStamp, timeZone, "dd.MM.YYYY HH:mm");
        var type = lastRowValues[0][1];
        var description = lastRowValues[0][2];
        var name = lastRowValues[0][3];
        var email = lastRowValues[0][4];
        var note = lastRowValues[0][5];
        var emailSubject = "Požadavek na podporu AppSatori - z " + shortTimeStamp + " JE POTŘEBA VAŠE REAKCE";
        var emailBody = "<i> U Vašeho požadavku s tímto popisem níže <strong>potřebujeme Vaši reakci</strong>\
        <p>" + description + "</p></i> \
        </p><p><hr /> \
        <strong>Co od Vás potřebujeme:</strong> \
        <p>" + note + "</p> \
        <p><i>S pozdravem</p> \
        Podpora AppSatori <br />\
        pomoc@appsatori.eu \
        ";
        
        //send email to the customer
        GmailApp.sendEmail(email,emailSubject,emailBody, {htmlBody: emailBody, replyTo: "pomoc@appsatori.eu"});
        
        //sets value "NO" to the "Notification" column
        notificationCol.setValue("NE").setVerticalAlignment("center");
       }
     }
    
    //3. condition - last update is older than today and state is "Done - Closed" and Notification doesn`t contain "NO"
    else if (dataday < today && specificState == "Hotovo - Uzavřeno" && notificationState !== "NE"){
      {
        var lastColumn = responsesSheet.getLastColumn();
        var lastRowValues = responsesSheet.getRange(i+1, 1, 7, lastColumn).getValues();
        Logger.log(lastRowValues)
        var timeZone = "GMT+2";
        var timeStamp = lastRowValues[0][0];
        var shortTimeStamp = Utilities.formatDate(timeStamp, timeZone, "dd.MM.YYYY HH:mm");
        var type = lastRowValues[0][1];
        var description = lastRowValues[0][2];
        var name = lastRowValues[0][3];
        var email = lastRowValues[0][4];
        var note = lastRowValues[0][5];
        var emailSubject = "Požadavek na podporu AppSatori - z "+ shortTimeStamp + " JE VYŘEŠEN";
        var emailBody = "<strong>Váš požadavek níže je vyřešen</strong>\
        <p>" + description + "</p> \
        </p><p><hr /> \
        <strong>S tímto výsledkem:</strong> \
        <p>" + note + "</p> \
        <p><i>S pozdravem</p> \
        Podpora AppSatori <br /> \
        pomoc@appsatori.eu \
        ";
        
        //send email to the customer
        GmailApp.sendEmail(email,emailSubject,emailBody, {htmlBody: emailBody, replyTo: "pomoc@appsatori.eu"});
        
        //sets value "NO" to the "Notification" column
        notificationCol.setValue("NE").setVerticalAlignment("center");
       }
     }
}
}
