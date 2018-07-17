Public Class UpdateForm
    Friend WithEvents TableLayout As TableLayoutPanel

    Public result ' stores the choice

    Private DataArray(,) As String ' stores the display data
    '    = New String(6, 1) {
    '    {"Full Name: Mary Ann Pacheco", "Full Name: Mary Ann Pacheco"},
    '    {"Company: Borough of Manhattan Community College", "Company: Rio Hondo Community College"},
    '    {"Job Title: Janitor", "Job Title: Trustee"},
    '    {"Email: fake@email.edu", "Email: fake@email.com"},
    '    {"Business Phone: 5556666", "Business Phone: 7775555"},
    '    {"Address: 1234 New Road Avenue" & vbNewLine & "CA  00002", "Address: 3333 Workman Mill Road" & vbNewLine & "PA  00001"},
    '    {"Notes: 2018 Position: Other 2018 Position: Other", ""}
    '}

    Public Sub New()
        CreateTable()
    End Sub

    Public Sub New(DataArray)
        Me.DataArray = DataArray

        CreateTable()
    End Sub

    ' creates the table for the DataArray
    Private Sub CreateTable()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AutoSizeMode = AutoSizeMode.GrowAndShrink
        AutoSize = True
        TableLayout = New TableLayoutPanel
        With TableLayout
            .Name = "tableLayout"
            .Margin = New Padding(0, 0, 0, 0)
            .ColumnCount = 0
            .RowCount = 0
            .Dock = DockStyle.Fill
            .AutoSizeMode = AutoSizeMode.GrowAndShrink
            .AutoScroll = True
        End With

        Controls.Add(TableLayout)
    End Sub

    Private Sub UpdateForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim title1 As Label = New Label
        With title1
            .Name = "Title1"
            .TextAlign = ContentAlignment.TopCenter
            .Text = "New Info:"
            .Dock = DockStyle.Left
            .AutoSize = True
        End With

        Dim title2 As Label = New Label
        With title2
            .Name = "Title2"
            .TextAlign = ContentAlignment.TopCenter
            .Text = "Old Info:"
            .Dock = DockStyle.Left
            .AutoSize = True
        End With

        Dim rowOffset As Integer = 1

        TableLayout.ColumnCount += 2
        TableLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        TableLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

        TableLayout.RowCount += 1
        TableLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        TableLayout.Controls.Add(title1, 0, 0)
        TableLayout.Controls.Add(title2, 1, 0)

        For x = LBound(DataArray, 1) To UBound(DataArray, 1) - 1
            'TableLayout.ColumnCount += 1
            'TableLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            For y = LBound(DataArray, 2) To UBound(DataArray, 2)
                If y = LBound(DataArray, 2) Then
                    TableLayout.RowCount += 1
                    TableLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                End If

                Dim lbl = New Label
                With lbl
                    .Name = "lbl" & x & y
                    .TextAlign = ContentAlignment.TopLeft
                    .Text = DataArray.GetValue(x, y)
                    .Dock = DockStyle.Left
                    .AutoSize = True
                End With

                TableLayout.Controls.Add(lbl, y, x + rowOffset)
            Next
        Next

        Dim notes As Label = New Label
    End Sub

    Private Sub Update_Click(sender As Object, e As EventArgs) Handles Update.Click
        result = Results.Update
        Close()
    End Sub

    Private Sub Submit_Click(sender As Object, e As EventArgs) Handles Submit.Click
        result = Results.Submit
        Close()
    End Sub

    Private Sub Skip_Click(sender As Object, e As EventArgs) Handles Skip.Click
        result = Results.Skip
        Close()
    End Sub

    Private Sub Notes_Click(sender As Object, e As EventArgs) Handles Notes.Click
        MsgBox(DataArray.GetValue(UBound(DataArray, 1), 0))
    End Sub
End Class