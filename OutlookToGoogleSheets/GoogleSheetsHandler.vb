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
    Private exportData As List(Of IList(Of Object))

    ' constructor
    Public Sub New()
        ApplicationName = "Outlook to Google Sheets"
        spreadsheetId = "insert spreadsheet ID here"
        exportData = New List(Of IList(Of Object))
    End Sub

    ' appends the export data list
    Public Sub AppendExportData(dataBlock As List(Of Object))
        exportData.Add(dataBlock)
    End Sub

    ' checks if exportData is empty
    Public Function IsEmpty()
        Return exportData.Count = 0
    End Function

    ' initializes communications with Google Sheets and submits the data
    Public Sub SubmitToGoogleSheets()
        Dim service = AuthorizeGoogleApp()
        Dim range As String = GetRange(service)

        'Dim requestValues As IList(Of IList(Of Object)) = BuildData()

        Dim requestbody As Data.ValueRange = New Data.ValueRange With {
            .Range = range,
            .MajorDimension = "1",
            .Values = exportData
        }

        UpdateGoogleSheetInBatch(requestbody, range, service)
    End Sub

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
    Private Sub BuildData()
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

        AppendExportData(obj)
    End Sub

    ' creates the request and submits the data to the
    Private Sub UpdateGoogleSheetInBatch(requestBody As Data.ValueRange, range As String, service As SheetsService)
        Dim request As SpreadsheetsResource.ValuesResource.AppendRequest = service.Spreadsheets.Values.Append(requestBody, spreadsheetId, range)
        request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
        Dim response As Data.AppendValuesResponse = request.Execute()
    End Sub
End Class