﻿'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Outlook to Google Sheets v1.0.2
' Created by Michael Cardenas ©2018
' 
' This application is used to gather contact information from e-mails 
' and store them as vcards within Outlook. The data that is gathered 
' in this process can also be submitted to a Google Sheets file 
' and/or saved as an Excel spreadhsheet.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Class MainForm
    Private Sub RunAndUpload_Click(sender As Object, e As EventArgs) Handles RunAndUpload.Click
        ErrorHandler.RunAndUpload()
    End Sub

    Private Sub RunMacro_Click(sender As Object, e As EventArgs) Handles RunMacro.Click
        ErrorHandler.RunMacro()
    End Sub

    Private Sub Upload_Click(sender As Object, e As EventArgs) Handles Upload.Click
        ErrorHandler.Upload()
    End Sub

    Private Sub Settings_Click(sender As Object, e As EventArgs) Handles Settings.Click
        Dim settings As Settings = New Settings()
        settings.ShowDialog()
    End Sub

    Private Sub Quit_Click(sender As Object, e As EventArgs) Handles Quit.Click
        Close()
    End Sub
End Class
