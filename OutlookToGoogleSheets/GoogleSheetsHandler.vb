'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GoogleSheetsHandler.vb
' Created by Michael Cardenas ©2018
' 
' This class handles the functionality required to communicated with
' Google Sheets.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Services
Imports Google.Apis.Util.Store
Imports System.IO
Imports System.Threading

Imports Data = Google.Apis.Sheets.v4.Data

Public Class GoogleSheetsHandler
    Private ApplicationName As String
    Private spreadsheetId As String
    Private sheetName As String
    Private column As String

    ' constructor
    Public Sub New()
        ApplicationName = "Outlook to Google Sheets"
        spreadsheetId = My.Settings.URL
        sheetName = My.Settings.SheetName
        column = My.Settings.Column
    End Sub

    ' initializes communications with Google Sheets and submits the data
    Public Sub SubmitToGoogleSheets(exportData As IList(Of IList(Of Object)))
        If exportData.Count = 0 Then
            Throw New Exception("There is currently no data to upload." & vbNewLine &
                "Please run the macro first.")
        End If

        Dim service = AuthorizeGoogleApp()
        Dim range As String = GetRange(service)

        Dim requestbody As Data.ValueRange = New Data.ValueRange With {
            .Range = range,
            .MajorDimension = "1",
            .Values = exportData
        }

        UpdateGoogleSheetInBatch(requestbody, range, service)
    End Sub

    Private Function Init()
        Dim fileReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader("init.txt")
        Dim stringReader As String = ""
        MsgBox("outside while loop")
        While fileReader.Peek() >= 0
            MsgBox("inside while loop")
            stringReader = fileReader.ReadLine()

            If stringReader.Contains("[SHEET ID]") Then
                stringReader.Replace("[SHEET ID]", "")
                stringReader.Replace(" ", "")
                Exit While
            End If
        End While

        Return stringReader
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

            ' save to log file
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
        'Dim range As String = "LogSheet!C:C"
        Dim range As String = sheetName & "!" & column & ":" & column
        Dim getRequest As SpreadsheetsResource.ValuesResource.GetRequest = service.Spreadsheets.Values.Get(spreadsheetId, range)
        Dim getResponse As Data.ValueRange = getRequest.Execute()
        Dim getValues As IList(Of IList(Of [Object])) = getResponse.Values
        Dim currentCount As Integer = getValues.Count() + 1

        Return sheetName & "!" & column & currentCount & ":" & column
    End Function

    ' used to generate data for testing purposes
    Private Function BuildData()
        Dim obj As List(Of Object) = New List(Of Object) From {
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

        Dim payload As List(Of IList(Of Object)) = New List(Of IList(Of Object))
        payload.Add(obj)

        Return payload
    End Function

    ' creates the request and submits the data to the Google Sheet
    Private Sub UpdateGoogleSheetInBatch(requestBody As Data.ValueRange, range As String, service As SheetsService)
        Dim request As SpreadsheetsResource.ValuesResource.AppendRequest = service.Spreadsheets.Values.Append(requestBody, spreadsheetId, range)
        request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
        Dim response As Data.AppendValuesResponse = request.Execute()
    End Sub
End Class