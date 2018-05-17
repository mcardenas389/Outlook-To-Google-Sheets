Public Class UpdateForm
    Friend WithEvents TableLayout As TableLayoutPanel
    Private DataArray(,) As String = New String(6, 1) {
        {"Full Name: Mary Ann Pacheco", "Full Name: Mary Ann Pacheco"},
        {"Company: Borough of Manhattan Community College", "Company: Rio Hondo Community College"},
        {"Job Title: Janitor", "Job Title: Trustee"},
        {"Email: angie.tomasich69@gmail.edu", "Email: angie.tomasich@riohondo.edu"},
        {"Business Phone: 5556969", "Business Phone: 5624637272"},
        {"Address: 3600 Workman Mill Road Whittier, CA  90601", "Address: 3600 Workman Mill Road Whittier, CA  90601"},
        {"Notes: 2018 Position: Other 2018 Position: Other", ""}
    }

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AutoSizeMode = AutoSizeMode.GrowAndShrink
        AutoSize = True
        TableLayout = New TableLayoutPanel
        With TableLayout
            .Name = "tableLayout"
            .Margin = New Padding(50, 50, 0, 0)
            .Location = New Point(12, 12)
            .ColumnCount = 0
            .RowCount = 0
            .Dock = DockStyle.Fill
            .AutoSizeMode = AutoSizeMode.GrowAndShrink
            .AutoSize = True
        End With

        Controls.Add(TableLayout)
    End Sub

    Private Sub UpdateForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        For x = LBound(DataArray, 1) To UBound(DataArray, 1)
            TableLayout.ColumnCount += 1
            TableLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            For y = LBound(DataArray, 2) To UBound(DataArray, 2)
                If y = LBound(DataArray, 2) Then
                    TableLayout.RowCount += 1
                    TableLayout.RowStyles.Add(New ColumnStyle(SizeType.AutoSize))
                End If

                Dim lbl = New Label
                With lbl
                    .Name = "lbl" & x & y
                    .TextAlign = ContentAlignment.TopLeft
                    .Text = DataArray.GetValue(x, y)
                    .Dock = DockStyle.Fill
                    .AutoSize = True
                End With

                TableLayout.Controls.Add(lbl, y, x)
            Next
        Next
    End Sub
End Class