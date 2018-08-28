# Outlook to Google Sheets

currently this readme is work in progress.

1. [Disclaimer](#disclaimer)
2. [Installation](#installation)
3. [Documentation](#documentation)

## Disclaimer

This application is only meant for internal use by the National Center for the Study of Collective Bargainning in Higher Education and the Professions.

## Installation

work in progress

## Documentation

On start-up you should see this window.

![main window](https://raw.githubusercontent.com/mcardenas389/Outlook-To-Google-Sheets/master/documentation%20images/main%20window.PNG)

From here you have a few options:

**Run Macro and Upload**<br />
This will allow you to run both the Outlook macro and also perform the upload to the Google Sheets roster. Both will be further explained in their respective sections.

**Run Macro**<br />
This will run the macro for Outlook. If Outlook is not already open, running the macro will also launch it for you. Here is a rundown of what to exepect.
  
![search folder window](https://raw.githubusercontent.com/mcardenas389/Outlook-To-Google-Sheets/master/documentation%20images/search%20folder%20window.PNG)
  
To ensure proper functionality, make sure that you have separated the e-mails you want to collect contact information from into their own folder. Simply type the name of that folder into the search bar and select OK to run the search.

![time frame window](https://raw.githubusercontent.com/mcardenas389/Outlook-To-Google-Sheets/master/documentation%20images/time%20frame%20window.png)

If the folder is found and has emails in it, you will be prompted to choose how far back you want to collect contact information from.

![update window](https://raw.githubusercontent.com/mcardenas389/Outlook-To-Google-Sheets/master/documentation%20images/update%20window.PNG)

If a contact does not exist in your default Outlook contacts folder, a contact will automatically be created and the application will not prompt you at all. However, if a contact is found to already exist, you'll be prompted to make a specific choice.

1. **Notes** - this will allow you to view the notes for this contact entry.
2. **Update** - this will update your contact information as well as send this information to the Google Sheets roster.
3. **Submit** - this will not update your contact information, but it still will send it to your Google Sheets roster.
4. **Skip** - this will ignore this entry and will neither update your contact information nor the roster.

Once you have gone through the batch, you will be prompted with how many emails were read and that the operation was a success.

**Upload**<br />
This is a purely autonamous operation and doesn't require any user input. It will simply upload the information that was collected when the macro was called. If the macro was not called beforehand or if it didn't collect any information, you will be notified that there is no information to send.

**Settings**<br />
![settings window](https://raw.githubusercontent.com/mcardenas389/Outlook-To-Google-Sheets/master/documentation%20images/settings%20window.PNG)

Here you can change a few settings of the application.

1. **Sheet ID**
2. **Sheet Name**
3. **Column**
