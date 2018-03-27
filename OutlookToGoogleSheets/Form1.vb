Public Class Form1
    Private Sub RunAndUpload_Click(sender As Object, e As EventArgs) Handles RunAndUpload.Click
        Module1.RunAndUpload()
    End Sub

    Private Sub RunMacro_Click(sender As Object, e As EventArgs) Handles RunMacro.Click
        Try
            Module1.RunMacro()
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Error")
        End Try
    End Sub

    Private Sub Upload_Click(sender As Object, e As EventArgs) Handles Upload.Click
        Module1.Upload()
    End Sub

    Private Sub Preview_Click(sender As Object, e As EventArgs) Handles Preview.Click
        'Module1.Preview()
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
