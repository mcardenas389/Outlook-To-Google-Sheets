﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class UpdateForm
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
        Me.Update = New System.Windows.Forms.Button()
        Me.Submit = New System.Windows.Forms.Button()
        Me.Skip = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Update
        '
        Me.Update.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Update.Location = New System.Drawing.Point(80, 180)
        Me.Update.Name = "Update"
        Me.Update.Size = New System.Drawing.Size(80, 23)
        Me.Update.TabIndex = 0
        Me.Update.Text = "&Update"
        Me.Update.UseVisualStyleBackColor = True
        '
        'Submit
        '
        Me.Submit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Submit.Location = New System.Drawing.Point(202, 180)
        Me.Submit.Name = "Submit"
        Me.Submit.Size = New System.Drawing.Size(80, 23)
        Me.Submit.TabIndex = 1
        Me.Submit.Text = "&Submit"
        Me.Submit.UseVisualStyleBackColor = True
        '
        'Skip
        '
        Me.Skip.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Skip.Location = New System.Drawing.Point(324, 180)
        Me.Skip.Name = "Skip"
        Me.Skip.Size = New System.Drawing.Size(80, 23)
        Me.Skip.TabIndex = 2
        Me.Skip.Text = "S&kip"
        Me.Skip.UseVisualStyleBackColor = True
        '
        'UpdateForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(488, 215)
        Me.Controls.Add(Me.Skip)
        Me.Controls.Add(Me.Submit)
        Me.Controls.Add(Me.Update)
        Me.MaximumSize = New System.Drawing.Size(620, 350)
        Me.Name = "UpdateForm"
        Me.Text = "Update?"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Update As Button
    Friend WithEvents Submit As Button
    Friend WithEvents Skip As Button
End Class