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

Public Class Form1
    Private errorHandler As ErrorHandler = New ErrorHandler()

    Private Sub RunAndUpload_Click(sender As Object, e As EventArgs) Handles RunAndUpload.Click
        errorHandler.RunAndUpload()
    End Sub

    Private Sub RunMacro_Click(sender As Object, e As EventArgs) Handles RunMacro.Click
        errorHandler.RunMacro()
    End Sub

    Private Sub Upload_Click(sender As Object, e As EventArgs) Handles Upload.Click
        errorHandler.Upload()
    End Sub

    Private Sub Preview_Click(sender As Object, e As EventArgs) Handles Preview.Click
        Form2.Show()
    End Sub

    Private Sub UploadFromFile_Click(sender As Object, e As EventArgs) Handles UploadFromFile.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
        End If
    End Sub

    Private Sub Quit_Click(sender As Object, e As EventArgs) Handles Quit.Click
        Close()
    End Sub
End Class
