Public Class ErrorHandler
    Private googleSheets As GoogleSheetsHandler
    Private bulkImport As BulkImportContacts

    Public Sub New()
        googleSheets = New GoogleSheetsHandler()
        bulkImport = New BulkImportContacts()
    End Sub

    ' runs the macro in Module2 and calls the functions necessary to upload data to Google Sheets
    Public Sub RunAndUpload()
        Try
            RunMacro()
            googleSheets.SubmitToGoogleSheets()
        Catch ex As Exception
            MsgBox(ex.Message, vbInformation, "Warning!")
        End Try
    End Sub

    ' checks if Outlook is running and then calls the macro defined in Module2
    ' throws an exception if Outlook is not found
    Public Sub RunMacro()
        Try
            bulkImport.Run()
        Catch ex As Exception
            MsgBox(ex.Message, vbInformation, "Warning!")
        End Try

        'Dim oApp As Outlook.Application = CheckForOutlook()
        Dim oApp As Outlook.Application = Nothing

        If oApp Is Nothing Then
            Throw New Exception("Outlook could not be found!")
        End If

        BulkImportContacts(oApp)
    End Sub

    ' 
    Public Sub Upload()
        If googleSheets.IsEmpty() Then
            Throw New Exception("There is currently no data to upload." & vbNewLine &
                "Please run the macro or load data from a file.")
        End If

        Try
            googleSheets.SubmitToGoogleSheets()
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Error!")
        End Try

        MsgBox("Process Completed!")
    End Sub

    Public Sub Preview()

    End Sub
End Class
