<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.RunAndUpload = New System.Windows.Forms.Button()
        Me.Quit = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.RunMacro = New System.Windows.Forms.Button()
        Me.Upload = New System.Windows.Forms.Button()
        Me.UploadFromFile = New System.Windows.Forms.Button()
        Me.Preview = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'RunAndUpload
        '
        Me.RunAndUpload.Location = New System.Drawing.Point(216, 12)
        Me.RunAndUpload.Name = "RunAndUpload"
        Me.RunAndUpload.Size = New System.Drawing.Size(93, 23)
        Me.RunAndUpload.TabIndex = 0
        Me.RunAndUpload.Text = "Run &and Upload"
        Me.RunAndUpload.UseVisualStyleBackColor = True
        '
        'Quit
        '
        Me.Quit.Location = New System.Drawing.Point(225, 157)
        Me.Quit.Name = "Quit"
        Me.Quit.Size = New System.Drawing.Size(75, 23)
        Me.Quit.TabIndex = 1
        Me.Quit.Text = "&Quit"
        Me.Quit.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.OutlookToGoogleSheets.My.Resources.Resources.O2GS_Logo
        Me.PictureBox1.Location = New System.Drawing.Point(12, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(170, 170)
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'RunMacro
        '
        Me.RunMacro.Location = New System.Drawing.Point(225, 41)
        Me.RunMacro.Name = "RunMacro"
        Me.RunMacro.Size = New System.Drawing.Size(75, 23)
        Me.RunMacro.TabIndex = 3
        Me.RunMacro.Text = "Run &Macro"
        Me.RunMacro.UseVisualStyleBackColor = True
        '
        'Upload
        '
        Me.Upload.Location = New System.Drawing.Point(225, 70)
        Me.Upload.Name = "Upload"
        Me.Upload.Size = New System.Drawing.Size(75, 23)
        Me.Upload.TabIndex = 4
        Me.Upload.Text = "&Upload"
        Me.Upload.UseVisualStyleBackColor = True
        '
        'UploadFromFile
        '
        Me.UploadFromFile.Location = New System.Drawing.Point(216, 99)
        Me.UploadFromFile.Name = "UploadFromFile"
        Me.UploadFromFile.Size = New System.Drawing.Size(93, 23)
        Me.UploadFromFile.TabIndex = 5
        Me.UploadFromFile.Text = "Upload from &File"
        Me.UploadFromFile.UseVisualStyleBackColor = True
        '
        'Preview
        '
        Me.Preview.Location = New System.Drawing.Point(225, 128)
        Me.Preview.Name = "Preview"
        Me.Preview.Size = New System.Drawing.Size(75, 23)
        Me.Preview.TabIndex = 6
        Me.Preview.Text = "&Preview"
        Me.Preview.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(341, 194)
        Me.Controls.Add(Me.Preview)
        Me.Controls.Add(Me.UploadFromFile)
        Me.Controls.Add(Me.Upload)
        Me.Controls.Add(Me.RunMacro)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Quit)
        Me.Controls.Add(Me.RunAndUpload)
        Me.Name = "Form1"
        Me.Text = "Outlook to Google Sheets"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents RunAndUpload As Button
    Friend WithEvents Quit As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents RunMacro As Button
    Friend WithEvents Upload As Button
    Friend WithEvents UploadFromFile As Button
    Friend WithEvents Preview As Button
End Class
