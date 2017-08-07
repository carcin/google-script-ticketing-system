//FUNCTION - ADD THE TICKET FROM THE FORM AND SEND EMAILS TO THE CUSTOMER AND SUPPORT
function onFormSubmit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responsesSheet = ss.getSheetByName("Odpovědi formuláře"); //rename the name of the answers sheet by your name
  
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
  var emailSubject = "Požadavek na podporu AppSatori - z "+ shortTimeStamp; //edit the subject of the email here to your language
  //edit your body of the email
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
  var copySubject = "Váš Požadavek na podporu AppSatori - z "+ shortTimeStamp; //edit the subject of the email here to your language
  //edit your body of the email
  var copyBody = "<h3><u>Ticket</u></h3> \
Toto je kopie Vašeho požadavku na podporu AppSatori. Potvrzujeme, že jsme jej obdrželi a do 24 hodin se Vám ozveme.<br /> \
<strong>Typ: </strong>"+ type +"<br /> \
<strong>Požadavek: </strong>"+ description +"<br /> \
</p><p><hr /> \
<strong>Od: </strong>"+ name +"<br /> \
<strong>Email: </strong>"+ email +"<br /> \
</p><p><hr /> \
<p>V případě, že jste tento požadavek nezasílali, kontaktujte nás prosím okamžitě.</p><br /> \
";

  //getting email adresses with support emails (in case of different support email addresses)
  var supportContacts = ss.getSheetByName("Emaily"); //edit the name of the sheet with your support emails sheet
  var numEmailRows = supportContacts.getLastRow();
  
  //condition for different support email adresses
  if(type==="Licence (přidání dalších licencí, prodloužení, odebrání...)") { //edit the condition to the value of your first selectbox type of ticket
    var emailTo = supportContacts.getRange(2, 2, numEmailRows, 1).getValues();
  } else {
    var emailTo = supportContacts.getRange(2, 1, numEmailRows, 1).getValues();
  }
  
  //sending emails to the support
  GmailApp.sendEmail(emailTo,emailSubject,emailBody, {htmlBody: emailBody, replyTo: "pomoc@appsatori.eu"}); //edit on your emails
  GmailApp.sendEmail(email,copySubject, copyBody, {htmlBody: copyBody, replyTo: "pomoc@appsatori.eu"}) //edit on your emails
}
  
//FUNCTION - ADD TODAY`S DATE IN CASE OF EDIT TICKET`S STATE
function onEdit(event) { 
  var timezone = "GMT+2"; //set your timezone
  var timestamp_format = "dd-MM-YYYY"; //set your timestamp format
  var updateColName = "Stav řešení"; //edit to the name of your column "G"
  var timeStampColName = "Datum posledního updatu"; //edit to the name of your column "H"
  var sheet = event.source.getSheetByName('Odpovědi formuláře'); //edit to the name of your sheet with the answeres

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responsesSheet = ss.getSheetByName("Odpovědi formuláře"); //edit to the name of your sheet with the answeres
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
    var lastUpdate = sheet.getRange(index, dateCol + 1); //Initialization of an active row and column "Last update date (H)"
    var notification = sheet.getRange(index, dateCol + 2); //Initialization of an active row and column "Notification (I)"
    var date = Utilities.formatDate(new Date(), "GMT+2", "dd.MM.YYYY HH:mm") //edit to your timezone and timeformat
    lastUpdate.setValue(date); //set today`s date in column "Last update date"
    notification.setValue("") //erase "NO" in column "Notification"
  }
}

//FUNCTION - COMPARE TODAY`S DATE AND DATE OF THE LAST UPDATE
function notification() {
  var responsesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Odpovědi formuláře"); //edit to the name of your sheet with the answeres
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
    if (dataday < today && specificState == "Řeší se" && notificationState !== "NE"){ //edit to your "In progress" value in the form and edit "NE" to "NO"
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
        notificationCol.setValue("NE").setVerticalAlignment("center"); //edit value to "NO"
      }
    }
    
    //2. condition - last update is older than today and state is "Waiting for the customer" and Notification doesn`t contain "NO"
    else if (dataday < today && specificState == "Čeká se na reakci zákazníka" && notificationState !== "NE"){ //edit to your "Waiting for the customer" value in the form and edit "NE" to "NO"
      {
        var lastColumn = responsesSheet.getLastColumn();
        var lastRowValues = responsesSheet.getRange(i+1, 1, 7, lastColumn).getValues();
        Logger.log(lastRowValues)
        var timeZone = "GMT+2"; //edit to your timezone
        var timeStamp = lastRowValues[0][0];
        var shortTimeStamp = Utilities.formatDate(timeStamp, timeZone, "dd.MM.YYYY HH:mm"); //edit to your timestamp
        var type = lastRowValues[0][1];
        var description = lastRowValues[0][2];
        var name = lastRowValues[0][3];
        var email = lastRowValues[0][4];
        var note = lastRowValues[0][5];
        var emailSubject = "Požadavek na podporu AppSatori - z " + shortTimeStamp + " JE POTŘEBA VAŠE REAKCE"; //edit to your subject
        //edit to your body of the email
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
        GmailApp.sendEmail(email,emailSubject,emailBody, {htmlBody: emailBody, replyTo: "pomoc@appsatori.eu"}); //change it to your email
        
        //sets value "NO" to the "Notification" column
        notificationCol.setValue("NE").setVerticalAlignment("center"); //edit value to "NO"
       }
     }
    
    //3. condition - last update is older than today and state is "Done - Closed" and Notification doesn`t contain "NO"
    else if (dataday < today && specificState == "Hotovo - Uzavřeno" && notificationState !== "NE"){ //edit to your "Done - Closed" value in the form and edit "NE" to "NO"
      {
        var lastColumn = responsesSheet.getLastColumn();
        var lastRowValues = responsesSheet.getRange(i+1, 1, 7, lastColumn).getValues();
        Logger.log(lastRowValues)
        var timeZone = "GMT+2"; //edit your timezone
        var timeStamp = lastRowValues[0][0];
        var shortTimeStamp = Utilities.formatDate(timeStamp, timeZone, "dd.MM.YYYY HH:mm"); //edit your timestamp
        var type = lastRowValues[0][1];
        var description = lastRowValues[0][2];
        var name = lastRowValues[0][3];
        var email = lastRowValues[0][4];
        var note = lastRowValues[0][5];
        var emailSubject = "Požadavek na podporu AppSatori - z "+ shortTimeStamp + " JE VYŘEŠEN"; //edit to your subject
        //edit to your body of the email
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
        GmailApp.sendEmail(email,emailSubject,emailBody, {htmlBody: emailBody, replyTo: "pomoc@appsatori.eu"}); //change it to your email
        
        //sets value "NO" to the "Notification" column
        notificationCol.setValue("NE").setVerticalAlignment("center"); //edit value to "NO"
       }
     }
}
}
