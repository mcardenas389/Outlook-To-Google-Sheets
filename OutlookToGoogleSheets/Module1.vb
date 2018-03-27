'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Outlook to Google Sheets
' Created by Michael Cardenas 2018
' 
' This application is used to gather contact information from e-mails 
' and store them as vcards within Outlook. The data that is gathered 
' in this process can also be submitted to a Google Sheets file 
' and/or saved as an Excel spreadhsheet.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Services
Imports Google.Apis.Util.Store
Imports System.IO
Imports System.Threading

Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Data = Google.Apis.Sheets.v4.Data

Module Module1
    Private ApplicationName = "Outlook to Google Sheets"
    Private spreadsheetId As [String] = "insert spreadsheet ID here"
    Private exportData As List(Of IList(Of Object)) = New List(Of IList(Of Object))

    ' runs the macro in Module2 and calls the functions necessary to upload data to Google Sheets
    Public Sub RunAndUpload()
        Try
            RunMacro()
            Upload()
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Error")
            Exit Sub
        End Try
    End Sub

    ' checks if Outlook is running and then calls the macro defined in Module2
    ' throws an exception if Outlook is not found
    Public Sub RunMacro()
        'Dim oApp As Outlook.Application = CheckForOutlook()
        Dim oApp As Outlook.Application = Nothing

        If oApp Is Nothing Then
            Throw New Exception("Outlook could not be found!")
            Exit Sub
        End If

        BulkImportContacts(oApp)
    End Sub

    ''''''''''change msgbox to an exception''''''''''
    Public Sub Upload()
        If IsEmpty() Then
            MsgBox("There is currently no data to upload." & vbNewLine &
                "Please run the macro or load data from a file.", vbInformation, "No Data")
            Exit Sub
        End If

        Dim service = AuthorizeGoogleApp()
        Dim range As String = GetRange(service)

        'Dim requestValues As IList(Of IList(Of Object)) = BuildData()

        Dim requestbody As Data.ValueRange = New Data.ValueRange With {
            .Range = range,
            .MajorDimension = "1",
            .Values = exportData
        }

        UpdateGoogleSheetInBatch(requestbody, range, service)

        MsgBox("Process Completed!")
    End Sub

    Public Sub Preview()

    End Sub

    ' stores data into exportData
    Public Sub AppendExportData(dataBlock As List(Of Object))
        exportData.Add(dataBlock)
    End Sub

    ' checks if exportData is empty
    Public Function IsEmpty()
        Return exportData.Count = 0
    End Function

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

    ' authorizes application to gain access to the Google Sheet
    Private Function AuthorizeGoogleApp()
        Dim credential As UserCredential
        Dim Scopes As String() = {SheetsService.Scope.Spreadsheets}

        ' send client_secret.json and store the credential
        Using stream = New FileStream("client_secret.json", FileMode.Open, FileAccess.Read)
            ' get global path and store credentials locally
            Dim credPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal)
            credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json")

            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                Scopes, "user", CancellationToken.None, New FileDataStore(credPath, True)).Result

            'Console.WriteLine(Convert.ToString("Credential file saved to: ") & credPath)
        End Using

        ' get service using the just obtained credentials
        Dim service = New SheetsService(New BaseClientService.Initializer() With {
            .HttpClientInitializer = credential,
            .ApplicationName = ApplicationName
        })

        Return service
    End Function

    ' finds the range where new entries can be submitted to the Google Sheet
    Private Function GetRange(service As SheetsService)
        'Define request parameters.
        Dim range As String = "Roster!A:A"
        Dim getRequest As SpreadsheetsResource.ValuesResource.GetRequest = service.Spreadsheets.Values.Get(spreadsheetId, range)
        Dim getResponse As Data.ValueRange = getRequest.Execute()
        Dim getValues As IList(Of IList(Of [Object])) = getResponse.Values
        Dim currentCount As Integer = getValues.Count() + 1

        Return "Roster!A" & currentCount & ":A"
    End Function

    ' used to generate data for testing purposes
    Private Function BuildData()
        Dim objNewRecords As List(Of IList(Of Object)) = New List(Of IList(Of Object))

        Dim obj As IList(Of Object) = New List(Of Object) From {
            "Column 1",
            "Column 2",
            "Column 3",
            "Column 4",
            "Column 5",
            "Column 6",
            "Column 7",
            "Column 8",
            "Column 9"
        }

        objNewRecords.Add(obj)

        Return objNewRecords
    End Function

    ' creates the request and submits the data to the
    Private Sub UpdateGoogleSheetInBatch(requestBody As Data.ValueRange, range As String, service As SheetsService)
        Dim request As SpreadsheetsResource.ValuesResource.AppendRequest = service.Spreadsheets.Values.Append(requestBody, spreadsheetId, range)
        request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
        Dim response As Data.AppendValuesResponse = request.Execute()
    End Sub
End Module