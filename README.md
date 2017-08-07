# Google Forms Ticketing

**This Google Script uses:**
## Google Forms
*for collecting tickets from the customer*

Question#1 - Type of the ticket *(Selectbox)*

Question#2 - Description of the request *(Long text field)*

Question#3 - Name of the customer *(Short text field)*

Question#4 - Email of the customer *(Short text field)*

## Google Spreadsheets
*for collecting answers from the customer*

Column#1 - Timestamp *(automatically filled from the form)*

Column#2 - Type of the ticket *(automatically filled from the form)*

Column#3 - Description of the request *(automatically filled from the form)*

Column#4 - Name of the customer *(automatically filled from the form)*

Column#5 - Email of the customer *(automatically filled from the form)*

Column#6 - Comment field *(When you want to send a custom message to the customer. More info in the script)*

Column#7 - Ticket state *(Automatically filled as "Open" by the script. You can change the state to "In progress", "Waiting for the customer" and "Done - Closed")*

Column#8 - Date of the last update *(Automatically filled with the current date and time. Serves for sending notification to the customer)*

Column#9 - Notification? *(Serves for tell the trigger not to send the notification email again. It sends just once when the state is changed and we don`t need to send it everytime)*

## Google Script
*for all the magic*
*to edit the Google Script click in the "Answers Spreadsheet" on Tools -> Script editor*

## How to prepare the Google Forms and Spreadsheet

1. Create new Google Form in Google Drive

2. Insert the questions mentioned above

3. Click on "Answers" and create a new Google Spreadsheet for the answers

4. Add the custom columns mentioned above

5. Create a new sheet for Emails where you can add different emails for different types of tickets *(more information in the script)*

6. Create another sheet for Dropdown menu for the States of the tickets *(I use: Open, In progress, Waiting for the customer and Done - Closed)*

7. Create Data validation for the column "G" with data from the Dropdown sheet

8. Open Script editor in Spreadsheet and insert the code

9. Rename the names of Sheets in the code and edit the HTMl of the emails
