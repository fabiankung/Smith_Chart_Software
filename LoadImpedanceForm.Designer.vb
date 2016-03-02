<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ZLForm
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
        Me.Input1TextBox = New System.Windows.Forms.TextBox()
        Me.Input2TextBox = New System.Windows.Forms.TextBox()
        Me.RLabel = New System.Windows.Forms.Label()
        Me.XLabel = New System.Windows.Forms.Label()
        Me.OkButton = New System.Windows.Forms.Button()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Input1TextBox
        '
        Me.Input1TextBox.Location = New System.Drawing.Point(48, 77)
        Me.Input1TextBox.Name = "Input1TextBox"
        Me.Input1TextBox.Size = New System.Drawing.Size(62, 20)
        Me.Input1TextBox.TabIndex = 0
        Me.Input1TextBox.Text = "50"
        '
        'Input2TextBox
        '
        Me.Input2TextBox.Location = New System.Drawing.Point(48, 103)
        Me.Input2TextBox.Name = "Input2TextBox"
        Me.Input2TextBox.Size = New System.Drawing.Size(62, 20)
        Me.Input2TextBox.TabIndex = 1
        Me.Input2TextBox.Text = "0"
        '
        'RLabel
        '
        Me.RLabel.AutoSize = True
        Me.RLabel.Location = New System.Drawing.Point(27, 80)
        Me.RLabel.Name = "RLabel"
        Me.RLabel.Size = New System.Drawing.Size(15, 13)
        Me.RLabel.TabIndex = 2
        Me.RLabel.Text = "R"
        '
        'XLabel
        '
        Me.XLabel.AutoSize = True
        Me.XLabel.Location = New System.Drawing.Point(27, 106)
        Me.XLabel.Name = "XLabel"
        Me.XLabel.Size = New System.Drawing.Size(14, 13)
        Me.XLabel.TabIndex = 3
        Me.XLabel.Text = "X"
        '
        'OkButton
        '
        Me.OkButton.Location = New System.Drawing.Point(129, 77)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.Size = New System.Drawing.Size(83, 37)
        Me.OkButton.TabIndex = 4
        Me.OkButton.Text = "OK"
        Me.OkButton.UseVisualStyleBackColor = True
        '
        'CancelButton
        '
        Me.CancelButton.Location = New System.Drawing.Point(129, 129)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(83, 37)
        Me.CancelButton.TabIndex = 5
        Me.CancelButton.Text = "Cancel"
        Me.CancelButton.UseVisualStyleBackColor = True
        '
        'ZLForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(245, 202)
        Me.Controls.Add(Me.CancelButton)
        Me.Controls.Add(Me.OkButton)
        Me.Controls.Add(Me.XLabel)
        Me.Controls.Add(Me.RLabel)
        Me.Controls.Add(Me.Input2TextBox)
        Me.Controls.Add(Me.Input1TextBox)
        Me.Name = "ZLForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Input1TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Input2TextBox As System.Windows.Forms.TextBox
    Friend WithEvents RLabel As System.Windows.Forms.Label
    Friend WithEvents XLabel As System.Windows.Forms.Label
    Friend WithEvents OkButton As System.Windows.Forms.Button
    Friend WithEvents CancelButton As System.Windows.Forms.Button
End Class
