﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form3
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.executeQuery = New System.Windows.Forms.Button()
        Me.serviceCall = New System.Windows.Forms.Button()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.ItemAliasText = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Connected to Database"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(12, 153)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(400, 96)
        Me.RichTextBox1.TabIndex = 4
        Me.RichTextBox1.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 134)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Result Box"
        '
        'executeQuery
        '
        Me.executeQuery.Location = New System.Drawing.Point(26, 83)
        Me.executeQuery.Name = "executeQuery"
        Me.executeQuery.Size = New System.Drawing.Size(97, 23)
        Me.executeQuery.TabIndex = 9
        Me.executeQuery.Text = "Execute Query"
        Me.executeQuery.UseVisualStyleBackColor = True
        '
        'serviceCall
        '
        Me.serviceCall.Location = New System.Drawing.Point(129, 83)
        Me.serviceCall.Name = "serviceCall"
        Me.serviceCall.Size = New System.Drawing.Size(75, 23)
        Me.serviceCall.TabIndex = 11
        Me.serviceCall.Text = "ServiceCall"
        Me.serviceCall.UseVisualStyleBackColor = True
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Location = New System.Drawing.Point(418, 48)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(323, 136)
        Me.RichTextBox2.TabIndex = 12
        Me.RichTextBox2.Text = ""
        '
        'ItemAliasText
        '
        Me.ItemAliasText.Location = New System.Drawing.Point(130, 48)
        Me.ItemAliasText.Name = "ItemAliasText"
        Me.ItemAliasText.Size = New System.Drawing.Size(100, 20)
        Me.ItemAliasText.TabIndex = 13
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(26, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(102, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Enter Barcode/Alias"
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(753, 261)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ItemAliasText)
        Me.Controls.Add(Me.RichTextBox2)
        Me.Controls.Add(Me.serviceCall)
        Me.Controls.Add(Me.executeQuery)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form3"
        Me.Text = "Form3"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents executeQuery As Button
    Friend WithEvents serviceCall As Button
    Friend WithEvents RichTextBox2 As RichTextBox
    Friend WithEvents ItemAliasText As TextBox
    Friend WithEvents Label3 As Label
End Class
