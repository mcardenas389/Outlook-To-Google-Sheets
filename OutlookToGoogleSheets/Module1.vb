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

    Public Sub Run()
        Dim oApp As Outlook.Application = CheckForOutlook()

        If oApp Is Nothing Then
            MsgBox("Outlook could not be found!", vbExclamation, "Error")
            Exit Sub
        End If

        BulkImportContacts(oApp)

        Dim service = AuthorizeGoogleApp()
        Dim range As String = GetRange(service)

        'Dim requestValues As IList(Of IList(Of Object)) = BuildData()

        Dim requestbody As Data.ValueRange = New Data.ValueRange With {
            .Range = range,
            .MajorDimension = "1",
            .Values = exportData
        }

        UpdateGoogleSheetInBatch(requestbody, spreadsheetId, range, service)

        Console.WriteLine("Complete!")
        Console.ReadKey()
    End Sub

    Private Function CheckForOutlook()
        Dim oApp As Outlook.Application = Nothing
        Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
            "Software\\microsoft\\windows\\currentversion\\app paths\\OUTLOOK.EXE")

        If key Is Nothing Then
            Return oApp
        End If

        Dim exePath As String = key.GetValue("Path")
        Dim processList() As Process = Process.GetProcessesByName("OUTLOOK")

        If Not exePath Is Nothing And processList.Length = 0 Then
            oApp = CreateObject("Outlook.Application")
            Process.Start(oApp.Name)
        ElseIf exePath Is Nothing Then
            Console.WriteLine("Outlook is not installed on this machine.")
        Else
            oApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application")
        End If

        Return oApp
    End Function

    Private Function AuthorizeGoogleApp()
        Dim credential As UserCredential
        Dim Scopes As String() = {SheetsService.Scope.Spreadsheets}

        Using stream = New FileStream("client_secret.json", FileMode.Open, FileAccess.Read)
            Dim credPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal)
            credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json")

            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                Scopes, "user", CancellationToken.None, New FileDataStore(credPath, True)).Result

            Console.WriteLine(Convert.ToString("Credential file saved to: ") & credPath)
        End Using

        Dim service = New SheetsService(New BaseClientService.Initializer() With {
            .HttpClientInitializer = credential,
            .ApplicationName = ApplicationName
        })

        Return service
    End Function

    Private Function GetRange(service As SheetsService)
        'Define request parameters.
        Dim range As String = "Sheet2!A:A"
        Dim getRequest As SpreadsheetsResource.ValuesResource.GetRequest = service.Spreadsheets.Values.Get(spreadsheetId, range)
        Dim getResponse As Data.ValueRange = getRequest.Execute()
        Dim getValues As IList(Of IList(Of [Object])) = getResponse.Values
        Dim currentCount As Integer = getValues.Count() + 1

        Return "Sheet1!A" & currentCount & ":A"
    End Function

    Private Function BuildData()
        Dim objNewRecords As List(Of IList(Of Object)) = New List(Of IList(Of Object))

        Dim obj As IList(Of Object) = New List(Of Object) From {
            "Column 1",
            "Column 2",
            "Column 3"
        }

        objNewRecords.Add(obj)

        Return objNewRecords
    End Function

    Private Sub UpdateGoogleSheetInBatch(requestBody As Data.ValueRange, spreadsheetId As String, range As String, service As SheetsService)
        Dim request As SpreadsheetsResource.ValuesResource.AppendRequest = service.Spreadsheets.Values.Append(requestBody, spreadsheetId, range)
        request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
        Dim response As Data.AppendValuesResponse = request.Execute()
    End Sub

    Public Sub AppendExportData(dataBlock As List(Of Object))
        exportData.Add(dataBlock)
    End Sub
End Module