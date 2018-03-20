Public Class Form1
    Private Sub RunAndUpload_Click(sender As Object, e As EventArgs) Handles RunAndUpload.Click
        Module1.RunAndUpload()
    End Sub

    Private Sub Quit_Click(sender As Object, e As EventArgs) Handles Quit.Click
        Close()
    End Sub
End Class
