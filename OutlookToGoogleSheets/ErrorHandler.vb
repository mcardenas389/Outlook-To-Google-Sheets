﻿'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ErrorHandler.vb
' Created by Michael Cardenas ©2018
' 
' This module handles some of the exceptions that may be thrown by the
' other classes that it calls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Module ErrorHandler
    Private bulkImport As BulkImportContacts

    ' constructor
    Sub New()
        bulkImport = New BulkImportContacts()
    End Sub

    ' runs the macro and calls the functions necessary to upload data to Google Sheets
    Public Sub RunAndUpload()
        Try
            bulkImport.Run()
            bulkImport.Upload()
            MsgBox("Upload Completed!")
        Catch ex As Exception
            MsgBox(ex.Message, vbInformation, "Warning!")
        End Try
    End Sub

    ' runs the macro
    Public Sub RunMacro()
        Try
            bulkImport.Run()
        Catch ex As Exception
            MsgBox(ex.Message, vbInformation, "Warning!")
        End Try
    End Sub

    ' uploads data gathered by BulkImportContacts() to the Google Sheet
    Public Sub Upload()
        Try
            bulkImport.Upload()
            MsgBox("Upload Completed!")
        Catch ex As Exception
            MsgBox(ex.Message, vbInformation, "Warning!")
        End Try
    End Sub
End Module