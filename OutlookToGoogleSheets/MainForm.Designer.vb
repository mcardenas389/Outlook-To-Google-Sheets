﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.RunAndUpload = New System.Windows.Forms.Button()
        Me.Quit = New System.Windows.Forms.Button()
        Me.RunMacro = New System.Windows.Forms.Button()
        Me.Upload = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Settings = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'RunAndUpload
        '
        Me.RunAndUpload.Location = New System.Drawing.Point(197, 30)
        Me.RunAndUpload.Name = "RunAndUpload"
        Me.RunAndUpload.Size = New System.Drawing.Size(93, 23)
        Me.RunAndUpload.TabIndex = 0
        Me.RunAndUpload.Text = "Run &and Upload"
        Me.RunAndUpload.UseVisualStyleBackColor = True
        '
        'Quit
        '
        Me.Quit.Location = New System.Drawing.Point(206, 146)
        Me.Quit.Name = "Quit"
        Me.Quit.Size = New System.Drawing.Size(75, 23)
        Me.Quit.TabIndex = 1
        Me.Quit.Text = "&Quit"
        Me.Quit.UseVisualStyleBackColor = True
        '
        'RunMacro
        '
        Me.RunMacro.Location = New System.Drawing.Point(206, 59)
        Me.RunMacro.Name = "RunMacro"
        Me.RunMacro.Size = New System.Drawing.Size(75, 23)
        Me.RunMacro.TabIndex = 3
        Me.RunMacro.Text = "Run &Macro"
        Me.RunMacro.UseVisualStyleBackColor = True
        '
        'Upload
        '
        Me.Upload.Location = New System.Drawing.Point(206, 88)
        Me.Upload.Name = "Upload"
        Me.Upload.Size = New System.Drawing.Size(75, 23)
        Me.Upload.TabIndex = 4
        Me.Upload.Text = "&Upload"
        Me.Upload.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(12, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(170, 170)
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'Settings
        '
        Me.Settings.Location = New System.Drawing.Point(206, 117)
        Me.Settings.Name = "Settings"
        Me.Settings.Size = New System.Drawing.Size(75, 23)
        Me.Settings.TabIndex = 5
        Me.Settings.Text = "Settings"
        Me.Settings.UseVisualStyleBackColor = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(308, 194)
        Me.Controls.Add(Me.Settings)
        Me.Controls.Add(Me.Upload)
        Me.Controls.Add(Me.RunMacro)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Quit)
        Me.Controls.Add(Me.RunAndUpload)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MainForm"
        Me.Text = "Outlook to Google Sheets v1.0.2"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents RunAndUpload As Button
    Friend WithEvents Quit As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents RunMacro As Button
    Friend WithEvents Upload As Button
    Friend WithEvents Settings As Button
End Class
