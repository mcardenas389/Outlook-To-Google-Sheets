Public Class Form1
    Private Sub Run_Click(sender As Object, e As EventArgs) Handles Run.Click
        Module1.Run()
    End Sub

    Private Sub Quit_Click(sender As Object, e As EventArgs) Handles Quit.Click
        Close()
    End Sub
End Class
