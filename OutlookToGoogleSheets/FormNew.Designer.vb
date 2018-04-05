<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormNew
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.RunAndUpload = New System.Windows.Forms.Button()
        Me.Quit = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.RunMacro = New System.Windows.Forms.Button()
        Me.Upload = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'RunAndUpload
        '
        Me.RunAndUpload.Location = New System.Drawing.Point(216, 30)
        Me.RunAndUpload.Name = "RunAndUpload"
        Me.RunAndUpload.Size = New System.Drawing.Size(93, 23)
        Me.RunAndUpload.TabIndex = 0
        Me.RunAndUpload.Text = "Run &and Upload"
        Me.RunAndUpload.UseVisualStyleBackColor = True
        '
        'Quit
        '
        Me.Quit.Location = New System.Drawing.Point(225, 117)
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
        Me.RunMacro.Location = New System.Drawing.Point(225, 59)
        Me.RunMacro.Name = "RunMacro"
        Me.RunMacro.Size = New System.Drawing.Size(75, 23)
        Me.RunMacro.TabIndex = 3
        Me.RunMacro.Text = "Run &Macro"
        Me.RunMacro.UseVisualStyleBackColor = True
        '
        'Upload
        '
        Me.Upload.Location = New System.Drawing.Point(225, 88)
        Me.Upload.Name = "Upload"
        Me.Upload.Size = New System.Drawing.Size(75, 23)
        Me.Upload.TabIndex = 4
        Me.Upload.Text = "&Upload"
        Me.Upload.UseVisualStyleBackColor = True
        '
        'FormNew
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(341, 194)
        Me.Controls.Add(Me.Upload)
        Me.Controls.Add(Me.RunMacro)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Quit)
        Me.Controls.Add(Me.RunAndUpload)
        Me.Name = "FormNew"
        Me.Text = "Outlook to Google Sheets"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents RunAndUpload As Button
    Friend WithEvents Quit As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents RunMacro As Button
    Friend WithEvents Upload As Button
End Class
