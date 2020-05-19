# Extract emails from a mailbox and export to an Excel sheet
Use de Gmail API with **nodejs**

## Description
Using de [Quickstart](https://developers.google.com/gmail/api/quickstart/nodejs) expone Gmail :

Using the API exposed by Gmail, I have implemented access to a mailbox that retrieves mail with a label and exports the most relevant information to an Excel sheet in this order:
1. I implemented a function (listLabels) to search for the ID of each mailbox label.
2. As a practice, I build the function (getLabel), where I get the name of the label and the number of emails tagged with it.
3. I look for the relation of messages, with the function (listMessages) in that mailbox for a specific label. This requires a pagination.
4. For each mail or message retrieved, I get the data of that message (showEachMessage) retrieved from the header (headers):
- **From**
- **To**
- **Date**
- **Subject**
In addition, I recover a first part of the body of the mail (snippet), knowing that I can exploit the first data of it that are marked as:
- **Nombre:**
- **Email:**
- **Teléfono:**
5. Finally, for each of the data indicated, a row is exported to an Excel sheet.

