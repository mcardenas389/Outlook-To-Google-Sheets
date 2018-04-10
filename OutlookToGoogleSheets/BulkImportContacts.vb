﻿'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BulkImportContacts.vb
' Created by Michael Cardenas 2018
' 
' This class handles getting the data from Outlook to save as contacts
' and also stores this information in a data structure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports Outlook = Microsoft.Office.Interop.Outlook

Public Class BulkImportContacts
    Private oApp As Outlook.Application
    Private exportData As List(Of IList(Of Object))
    Private gSheets As GoogleSheetsHandler

    Public Sub New()
        exportData = New List(Of IList(Of Object))
        gSheets = New GoogleSheetsHandler()
    End Sub

    ' send upload data to the Google sheet
    Public Sub Upload()
        gSheets.SubmitToGoogleSheets(exportData)
    End Sub

    ' imports contacts from emails from a folder
    ' searches for the folder with FindInFolders()
    ' calls ImportToContacts() when folder is found
    Public Sub Run()
        oApp = CheckForOutlook()

        If oApp Is Nothing Then
            Throw New Exception("Outlook could not be found!")
        End If

        Dim name As String

        ''''''remove after testing''''''
        Dim nSpace As Outlook.NameSpace = oApp.GetNamespace("MAPI")
        Dim folder As Outlook.Folder = nSpace.Folders("ncscbint2@hunter.cuny.edu").Folders("Inbox").Folders("Test")
        ImportToContacts(folder)
        Exit Sub

        ' if nothing is entered, exit out of macro
        name = InputBox("Enter ConstantContact folder name:", "Search Folder")
        If Len(Trim$(name)) = 0 Then Exit Sub

        ' find out if folder exists
        Dim FoundFolder As Outlook.Folder
        FoundFolder = FindInFolders(oApp.Session.Folders, name)

        ' if folder is not found or is empty, do nothing
        ' if the folder is found and has items, call ImportToContacts()
        If FoundFolder Is Nothing Then
            MsgBox("Folder not found.", vbInformation)
            Exit Sub
        ElseIf FoundFolder.Items.Count = 0 Then
            MsgBox("Folder is empty.")
            Exit Sub
        Else
            Call ImportToContacts(FoundFolder)
        End If

        ' clean up
        FoundFolder = Nothing
    End Sub

    ' checks if Outlook is installed on the machine.
    ' returns Nothing if it is not.
    ' returns an instance of Outlook if it is.
    Private Function CheckForOutlook()
        Dim oApp As Outlook.Application = Nothing

        ' find Outlook in its default path
        Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
            "Software\\microsoft\\windows\\currentversion\\app paths\\OUTLOOK.EXE")

        If key Is Nothing Then
            Return oApp
        End If

        Dim exePath As String = key.GetValue("Path")

        ' check if Outlook is already running
        Dim processList() As Process = Process.GetProcessesByName("OUTLOOK")

        ' if Outlook is not running, launch it and return the instance
        ' if Outlook is running, get and return its instance
        If Not exePath Is Nothing And processList.Length = 0 Then
            oApp = CreateObject("Outlook.Application")
            Process.Start(oApp.Name)
        ElseIf exePath Is Nothing Then
            MsgBox("Outlook is not installed on this machine.", vbExclamation, "Outlook Not Found")
        Else
            oApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application")
        End If

        Return oApp
    End Function

    ' searches for a given folder name in the inbox
    Private Function FindInFolders(TheFolders As Outlook.Folders, name As String)
        Dim SubFolder As Outlook.Folder

        On Error Resume Next

        FindInFolders = Nothing

        For Each SubFolder In TheFolders
            If LCase(SubFolder.Name) Like LCase(name) Then
                ' return value
                FindInFolders = SubFolder
                Exit For
            Else
                ' return value
                FindInFolders = FindInFolders(SubFolder.Folders, name)
                If Not FindInFolders Is Nothing Then Exit For
            End If
        Next
    End Function

    ' searches through a given folder and gets contact information from email body
    ' calls CreateOrUpdateContact() to create new contacts
    Private Sub ImportToContacts(FoundFolder As Outlook.Folder)
        Dim Mail As Outlook.MailItem
        Dim MyItems As Outlook.Items
        Dim choice As String = ""
        Dim constantContact As String = "donotreply_eventspot@constantcontact.com"
        Dim paypal As String = "service@paypal.com"
        Dim filter As String
        Dim counter As Integer = 0

        On Error GoTo ErrorHandler

        Dim timeFrame As TimeFrame = New TimeFrame()

        Do While choice Is ""
            If timeFrame.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                choice = timeFrame.TimeFrames.Text
            Else
                Exit Sub
            End If
        Loop

        Dim days As Integer = 0

        ' set days if choice is not Today
        If choice = "Yesterday" Then
            days = 1
        ElseIf choice = "A week" Then
            days = 7
        ElseIf choice = "Two weeks" Then
            days = 14
        ElseIf choice = "30 days" Then
            days = 30
        End If

        filter = "[Received] >= " & Chr(34) & New DateTime(2017, 8, 30) & Chr(34)
        'filter = "[Received] >= " & Chr(34) & DateTime.Today.AddDays(-days) & Chr(34)

        MyItems = FoundFolder.Items.Restrict(filter)

        For Each Mail In MyItems
            ' mark as read if it is unread
            If Mail.UnRead Then
                Mail.UnRead = False
            End If

            ' If Mail.SenderEmailAddress Like constantContact Then
            Call CreateOrUpdateContact(Mail.Body)
            ' ElseIf Mail.SenderEmailAddress Like paypal Then
            '    Call UpdatePayment(Mail.body)
            ' End If

            ' for debugging
            ' Mail.SaveAs "C:\Users\Hunter\Documents\out" & counter & ".txt", olTXT
            ' Mail.SaveAs "C:\Users\Michelle\Documents\out" & counter & ".txt", olTXT
            ' Debug.Print counter + 1 & ". Body: " & vbNewLine & Mail.body

            counter += 1
        Next

        Dim msg As String = "Process was successful!" & vbNewLine & vbNewLine &
            "Read " & counter & " e-mails."

        MsgBox(msg, , "Success!")

ErrorHandler:
        If Err.Number <> 0 Then
            msg = "Error # " & Str(Err.Number) & " was generated by " &
                Err.Source & Chr(13) & Err.Description
            MsgBox(msg, vbExclamation, "Error!")
        End If

        ' clean up
        MyItems = Nothing
        Mail = Nothing
    End Sub

    ' update the notes of a given contact to display that payment was received
    Private Sub UpdatePayment(body As String)
        Dim messageArray() As String
        Dim splitArray() As String
        Dim delimitedMessage As String
        Dim fullName1 As String
        Dim fullName2 As String
        Dim email As String
        Dim payment As String
        Dim prompt As String
        Dim Contact As Outlook.ContactItem
        Dim ContactItems As Outlook.Items

        ' replace specific text with ### in order to split it up into an array
        ' field names may change if email body changes in the future
        delimitedMessage = Replace(body, "Buyer information", "###")
        delimitedMessage = Replace(delimitedMessage, "Instructions from buyer", "###")
        delimitedMessage = Replace(delimitedMessage, "National Center's Regional Conference - ", "###")
        delimitedMessage = Replace(delimitedMessage, "Insurance:", "###")
        messageArray = Split(delimitedMessage, "###")

        ' clean up fullName1
        splitArray = Split(messageArray(1), vbNewLine)
        fullName1 = splitArray(1)
        email = splitArray(2)

        ' clean up email
        splitArray = Split(email, Chr(34))
        email = splitArray(UBound(splitArray))

        ' clean up fullName2
        splitArray = Split(messageArray(3), vbNewLine)
        fullName2 = splitArray(0)

        ' clean up payment
        splitArray = Split(splitArray(1), " ")
        payment = Replace(splitArray(0), vbTab, "")

        splitArray = Split(fullName2, " ")
        ContactItems = FindContacts(splitArray(0), splitArray(1))

        ' prompt the user if there are multiple results from the query
        If ContactItems.Count > 1 Then
            Dim msg As String = "There are " & ContactItems.Count & " matches for " & Chr(34) & fullName2 & Chr(34) & "!"
            MsgBox(msg, , "Paypal: Multiple Results!")
        End If

        For Each Contact In ContactItems
            prompt = "Name: " & Contact.FullName & vbNewLine &
            "Email: " & Contact.Email1Address & vbNewLine &
            "Phone: " & Contact.BusinessTelephoneNumber & vbNewLine &
            "Company: " & Contact.CompanyName & vbNewLine &
            "Job Title: " & Contact.JobTitle & vbNewLine &
            "Address: " & Contact.BusinessAddress & vbNewLine &
            Contact.BusinessAddressCountry & vbNewLine &
            "Notes: " & Contact.Body & vbNewLine &
            vbNewLine & "Update with pyament information?"

            If MsgBox(prompt, vbQuestion Or vbYesNo, "Update?") = vbYes Then
                ' adjust for paypal fee
                Dim value As Double
                value = CDec(payment) - (CDec(payment) * 0.022 + 0.3)

                splitArray = Split(Contact.Body, "Total payment:")

                If UBound(splitArray) = 0 Then
                    Contact.Body = Contact.Body & vbNewLine & "Total payment: $" & value

                    If fullName1 <> fullName2 Then
                        Contact.Body = Contact.Body & " c/o " & fullName1
                    End If

                    Contact.Body = Contact.Body & " " '& Year(Of Date)()
                    ' Contact.Save
                End If
            End If
        Next

        '    If Not Contact Is Nothing Then
        '
        '    Else
        '        msg = "No entry found for " & Chr(34) & fullName2 & Chr(34) & "!"
        '        MsgBox msg, vbExclamation, "Paypal Update"
        '    End If

        ' clean up
        Contact = Nothing
    End Sub

    ' gets contact information from email body and uses this information to populate a contact card
    Private Sub CreateOrUpdateContact(body As String)
        Dim messageArray() As String
        Dim splitArray() As String
        Dim delimitedMessage As String
        Dim prompt As String
        Dim ContactItems As Outlook.Items
        Dim Contact As Outlook.ContactItem

        On Error GoTo ErrorHandler

        ' replace specific text with ### in order to split it up into an array
        ' field names may change if email body changes in the future
        delimitedMessage = Replace(body, "First Name:", "###")
        delimitedMessage = Replace(delimitedMessage, "Last Name:", "###")
        delimitedMessage = Replace(delimitedMessage, "Email Address:", "###")
        delimitedMessage = Replace(delimitedMessage, "Phone:", "###")
        delimitedMessage = Replace(delimitedMessage, "Business Information", "###")
        delimitedMessage = Replace(delimitedMessage, "Company:", "###")
        delimitedMessage = Replace(delimitedMessage, "Job Title:", "###")
        delimitedMessage = Replace(delimitedMessage, "Address 1:", "###")
        delimitedMessage = Replace(delimitedMessage, "City:", "###")
        delimitedMessage = Replace(delimitedMessage, "State:", "###")
        delimitedMessage = Replace(delimitedMessage, "ZIP Code:", "###")
        delimitedMessage = Replace(delimitedMessage, "Country:", "###")
        delimitedMessage = Replace(delimitedMessage, "What is your position?", "###")
        delimitedMessage = Replace(delimitedMessage, "Payment Summary", "###")
        delimitedMessage = Replace(delimitedMessage, "Total", "###")
        messageArray = Split(delimitedMessage, "###")

        ' clean up values and remove unwanted characters
        ' used on shared PC
        Dim i As Integer
        For i = 1 To 13
            ' remove the " mark from the hyperlink
            If i = 3 Or i = 4 Or i = 8 Then
                splitArray = Split(messageArray(i), Chr(34))
                messageArray(i) = splitArray(UBound(splitArray))
            End If

            ' remove the newline character and replace it with an empty string
            messageArray(i) = Replace(messageArray(i), vbNewLine, "")
        Next

        splitArray = Split(messageArray(15), vbNewLine)
        messageArray(15) = splitArray(6)

        '' replace unwanted characters with an empty string
        '' used on end user's PC
        'Dim i As Integer
        'For i = 1 To UBound(messageArray)
        '    messageArray(i) = Replace(messageArray(i), vbNewLine, "")
        '    messageArray(i) = Replace(messageArray(i), vbTab, "")
        'Next

        ' search for contacts after collecting the relevant data
        ContactItems = FindContacts(messageArray(1), messageArray(2))

        ' prompt the user if there are multiple results from the query
        If ContactItems.Count > 1 Then
            Dim msg As String = "There are " & ContactItems.Count & " matches for " & Chr(34) &
            messageArray(1) & " " & messageArray(2) & Chr(34) & "!"
            MsgBox(msg, , "ConstantContact: Multiple Results!")
        End If

        ' if the contact is found, then prompt the user before updating it
        For Each Contact In ContactItems
            ' build prompt
            ' new contact info
            prompt = "Contact exists!" & vbNewLine & vbNewLine & "New information:" & vbNewLine &
                "Name: " & messageArray(1) & " " & messageArray(2) & vbNewLine &
                "Email: " & messageArray(3) & vbNewLine &
                "Phone: " & messageArray(4) & vbNewLine &
                "Company: " & messageArray(6) & vbNewLine &
                "Job Title: " & messageArray(7) & vbNewLine &
                "Address: " & messageArray(8) & vbNewLine & messageArray(9) & ", " &
                StateAbbreviation(messageArray(10)) & " " & messageArray(11) & vbNewLine &
                messageArray(12) & vbNewLine & vbNewLine

            ' old contact info
            prompt = prompt & "Old information:" & vbNewLine &
                "Name: " & Contact.FullName & vbNewLine &
                "Email: " & Contact.Email1Address & vbNewLine &
                "Phone: " & Contact.BusinessTelephoneNumber & vbNewLine &
                "Company: " & Contact.CompanyName & vbNewLine &
                "Job Title: " & Contact.JobTitle & vbNewLine &
                "Address: " & Contact.BusinessAddress & vbNewLine &
                Contact.BusinessAddressCountry & vbNewLine &
                "Notes: " & Contact.Body & vbNewLine &
                vbNewLine & "Update with new information?"

            If MsgBox(prompt, vbQuestion Or vbYesNo, "Update?") = vbNo Then
                Contact = Nothing
            End If

            ' create or update contact if contact object has been set
            Call SaveContact(Contact, messageArray)
        Next

        ' if no contacts are found, then create a new contact without prompting the user
        If ContactItems.Count = 0 Then
            Call SaveContact(oApp.CreateItem(Outlook.OlItemType.olContactItem), messageArray)
        End If

ErrorHandler:
        If Err.Number <> 0 Then
            Dim msg As String = "Error # " & Str(Err.Number) & " was generated by " &
            Err.Source & Chr(13) & Err.Description
            MsgBox(msg, , "Error!") ', Err.HelpFile, Err.HelpContext)
            End
        End If

        ' clean up
        Contact = Nothing
        ContactItems = Nothing
    End Sub

    ' using a given firstName and lastName searches and returns a collection of Contacts
    Private Function FindContacts(firstName As String, lastName As String)
        Dim filter As String
        Dim ContactsFolder As Outlook.Folder
        Dim ContactItems As Outlook.Items

        filter = "[FullName] = " & Chr(34) & firstName & " " & lastName & Chr(34)

        ContactsFolder = oApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
        ContactItems = ContactsFolder.Items.Restrict(filter)

        ' return value
        FindContacts = ContactItems

        ' clean up
        ContactsFolder = Nothing
        ContactItems = Nothing
    End Function

    ' searches for a contact using a given first name, last name, and email
    ' returns contact object if it is found and Nothing if it is not found
    Private Function FindContact(firstName As String, lastName As String, Optional email As String = "email@default.com")
        Dim filter As String
        Dim ContactsFolder As Outlook.Folder
        Dim Contact As Outlook.ContactItem

        filter = "[FullName] = " & Chr(34) & firstName & " " & lastName & Chr(34)

        If email <> "" Then
            filter = filter & " And [E-mail] = " & Chr(34) & email & Chr(34)
        End If

        ContactsFolder = oApp.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
        Contact = ContactsFolder.Items.Find(filter)

        ' return value
        FindContact = Contact

        ' clean up
        ContactsFolder = Nothing
        Contact = Nothing
    End Function

    ' saves contacts into the default contact folder of Outlook
    Private Sub SaveContact(Contact As Outlook.ContactItem, messageArray() As String)
        ' check to see if an empty object was passed through
        If Not Contact Is Nothing Then
            With Contact
                .FirstName = messageArray(1)
                .LastName = messageArray(2)
                .Email1Address = messageArray(3)
                .BusinessTelephoneNumber = messageArray(4)
                .CompanyName = messageArray(6)
                .JobTitle = messageArray(7)
                .BusinessAddressStreet = messageArray(8)
                .BusinessAddressCity = messageArray(9)
                .BusinessAddressState = StateAbbreviation(messageArray(10))
                .BusinessAddressPostalCode = messageArray(11)
                .BusinessAddressCountry = messageArray(12)
                .Categories = "Correspondence"
            End With

            If Contact.Body = "" Then
                Contact.Body = DateTime.Today.Year & vbNewLine & "Position: " & messageArray(13)
            Else
                Dim prompt As String = "Append notes?" & vbNewLine & vbNewLine & "Notes:" & vbNewLine &
                    Contact.Body & vbNewLine & vbNewLine & "Append with:" & vbNewLine &
                    "Position: " & messageArray(13)

                If MsgBox(prompt, vbQuestion Or vbYesNo) = vbYes Then
                    Contact.Body = Contact.Body & vbNewLine & vbNewLine & DateTime.Today.Year &
                        vbNewLine & "Position: " & messageArray(13)
                End If
            End If

            ' build export data
            Call BuildExportData(messageArray)

            ' save contact data in default contacts folder
            Contact.Save()
        End If
    End Sub

    ' creates a list of objects containing contact data
    ' submits it to Module1.AppendExportData() to become of data payload
    Private Sub BuildExportData(messageArray() As String)
        Dim dataBlock As List(Of Object) = New List(Of Object) From {
            messageArray(2),
            messageArray(1),
            messageArray(7),
            messageArray(6),
            messageArray(9),
            StateAbbreviation(messageArray(10)),
            messageArray(12),
            messageArray(4),
            messageArray(3)
        }

        exportData.Add(dataBlock)
    End Sub

    ' returns the state abbreviation
    ' returns the original string if it is not found
    Private Function StateAbbreviation(stateName As String)
        Dim sn As String
        sn = UCase(stateName)

        StateAbbreviation = stateName

        If sn = "ALABAMA" Then
            StateAbbreviation = "AL"
        ElseIf sn = "ALASKA" Then
            StateAbbreviation = "AK"
        ElseIf sn = "ARIZONA" Then
            StateAbbreviation = "AZ"
        ElseIf sn = "ARKANSAS" Then
            StateAbbreviation = "AR"
        ElseIf sn = "CALIFORNIA" Then
            StateAbbreviation = "CA"
        ElseIf sn = "COLORADO" Then
            StateAbbreviation = "CO"
        ElseIf sn = "CONNECTICUT" Then
            StateAbbreviation = "CT"
        ElseIf sn = "DELAWARE" Then
            StateAbbreviation = "DE"
        ElseIf sn = "FLORIDA" Then
            StateAbbreviation = "FL"
        ElseIf sn = "GEORGIA" Then
            StateAbbreviation = "GA"
        ElseIf sn = "HAWAII" Then
            StateAbbreviation = "HI"
        ElseIf sn = "IDAHO" Then
            StateAbbreviation = "ID"
        ElseIf sn = "ILLINOIS" Then
            StateAbbreviation = "IL"
        ElseIf sn = "INDIANA" Then
            StateAbbreviation = "IN"
        ElseIf sn = "IOWA" Then
            StateAbbreviation = "IA"
        ElseIf sn = "KANSAS" Then
            StateAbbreviation = "KS"
        ElseIf sn = "KENTUCKY" Then
            StateAbbreviation = "KY"
        ElseIf sn = "LOUISIANA" Then
            StateAbbreviation = "LA"
        ElseIf sn = "MAINE" Then
            StateAbbreviation = "ME"
        ElseIf sn = "MARYLAND" Then
            StateAbbreviation = "MD"
        ElseIf sn = "MASSACHUSETTS" Then
            StateAbbreviation = "MA"
        ElseIf sn = "MICHIGAN" Then
            StateAbbreviation = "MI"
        ElseIf sn = "MINNESOTA" Then
            StateAbbreviation = "MN"
        ElseIf sn = "MISSISSIPPI" Then
            StateAbbreviation = "MS"
        ElseIf sn = "MISSOURI" Then
            StateAbbreviation = "MO"
        ElseIf sn = "MONTANA" Then
            StateAbbreviation = "MT"
        ElseIf sn = "NEBRASKA" Then
            StateAbbreviation = "NE"
        ElseIf sn = "NEVADA" Then
            StateAbbreviation = "NV"
        ElseIf sn = "NEW HAMPSHIRE" Then
            StateAbbreviation = "NH"
        ElseIf sn = "NEW JERSEY" Then
            StateAbbreviation = "NJ"
        ElseIf sn = "NEW MEXICO" Then
            StateAbbreviation = "NM"
        ElseIf sn = "NEW YORK" Then
            StateAbbreviation = "NY"
        ElseIf sn = "NORTH CAROLINA" Then
            StateAbbreviation = "NC"
        ElseIf sn = "NORTH DAKOTA" Then
            StateAbbreviation = "ND"
        ElseIf sn = "OHIO" Then
            StateAbbreviation = "OH"
        ElseIf sn = "OKLAHOMA" Then
            StateAbbreviation = "OK"
        ElseIf sn = "OREGON" Then
            StateAbbreviation = "OR"
        ElseIf sn = "PENNSYLVANIA" Then
            StateAbbreviation = "PA"
        ElseIf sn = "RHODE ISLAND" Then
            StateAbbreviation = "RI"
        ElseIf sn = "SOUTH CAROLINA" Then
            StateAbbreviation = "SC"
        ElseIf sn = "SOUTH DAKOTA" Then
            StateAbbreviation = "SD"
        ElseIf sn = "TENNESSEE" Then
            StateAbbreviation = "TN"
        ElseIf sn = "TEXAS" Then
            StateAbbreviation = "TX"
        ElseIf sn = "UTAH" Then
            StateAbbreviation = "UT"
        ElseIf sn = "VERMONT" Then
            StateAbbreviation = "VT"
        ElseIf sn = "VIRGINIA" Then
            StateAbbreviation = "VA"
        ElseIf sn = "WASHINGTON" Then
            StateAbbreviation = "WA"
        ElseIf sn = "WEST VIRGINIA" Then
            StateAbbreviation = "WV"
        ElseIf sn = "WISCONSIN" Then
            StateAbbreviation = "WI"
        ElseIf sn = "WYOMING" Then
            StateAbbreviation = "WY"
        End If
    End Function
End Class