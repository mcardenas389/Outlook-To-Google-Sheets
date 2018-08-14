Public Class Settings
    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextBox1.Text = My.Settings.URL
        TextBox2.Text = My.Settings.SheetName
        TextBox3.Text = My.Settings.Column
    End Sub

    Private Sub Save_Click(sender As Object, e As EventArgs) Handles Save.Click
        My.Settings.URL = TextBox1.Text
        My.Settings.SheetName = TextBox2.Text
        My.Settings.Column = TextBox3.Text
        Close()
    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Close()
    End Sub
End Class