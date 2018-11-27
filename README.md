# Outlook to Google Sheets

1. [Disclaimer](#disclaimer)
2. [Installation](#installation)
3. [Documentation](#documentation)
4. [Version History](#version-history)

## Disclaimer

This application is only meant for internal use by the National Center for the Study of Collective Bargainning in Higher Education and the Professions.

## Installation

To download the latest release of this project, simply follow this [link](https://github.com/mcardenas389/Outlook-To-Google-Sheets/releases/tag/v1.0) then download and unzip the file.

However, you won't be able to upload any data you collect from Outlook just yet. For the sake of security, the client_secret.json file is not included in the zip file. All you need to do is download the json file from [Google APIs](https://console.developers.google.com), go to APIs & Services, and click on Credentials. Click on the download button from the revelant . Once you have the json file, **remember to rename it to client_secret.json** and include it in the root folder of the project. You should now be able to upload data to your desired Google Sheets spreadsheet.

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

1. **Sheet ID** - this can be located within the URL of your Google Sheets spreadsheet. It is a long sequence of random alphanumeric characters. Just copy that and paste it here should you change your spreadsheet in the future.
2. **Sheet Name** - this is the name of the sheet within your spreadsheet you want to store the information in.
3. **Column** - this is where you'll start writing information. The information is always written on a single row.

**Quit**<br />
This will close out the application. If you have neglected to upload your information, you will have to run the macro again on start-up. This application does not save the information you want to send to Google Sheets.

## Version History
**v1.0.1**<br />

* Changed the order the columns are written in. Now First Name comes first.

* Fixed it so that the notes for the contact are updated for both Update and Submit options.

**v1.0**<br />

* Initial release.
