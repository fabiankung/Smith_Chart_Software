Public Class OptionForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents ZoTextBox As System.Windows.Forms.TextBox
    Friend WithEvents FreqTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ZoTextBox = New System.Windows.Forms.TextBox
        Me.FreqTextBox = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.OKButton = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'ZoTextBox
        '
        Me.ZoTextBox.Location = New System.Drawing.Point(112, 40)
        Me.ZoTextBox.Name = "ZoTextBox"
        Me.ZoTextBox.Size = New System.Drawing.Size(96, 20)
        Me.ZoTextBox.TabIndex = 0
        Me.ZoTextBox.Text = ""
        '
        'FreqTextBox
        '
        Me.FreqTextBox.Location = New System.Drawing.Point(112, 104)
        Me.FreqTextBox.Name = "FreqTextBox"
        Me.FreqTextBox.Size = New System.Drawing.Size(96, 20)
        Me.FreqTextBox.TabIndex = 1
        Me.FreqTextBox.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 24)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Reference Impedance Zo"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 32)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Operating frequency"
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(96, 144)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(64, 32)
        Me.OKButton.TabIndex = 4
        Me.OKButton.Text = "OK"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(216, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Ohm"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(216, 104)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 24)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "MHz"
        '
        'OptionForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(264, 190)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.FreqTextBox)
        Me.Controls.Add(Me.ZoTextBox)
        Me.Name = "OptionForm"
        Me.Text = "OptionForm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Module level declarations
    Public mstrFormTitle As String
    Public mstrVersion As String
    Public InterFormMessage As New MessageClass

    Private Sub OptionForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            'Update the title and version of the main window
            Me.Text = mstrFormTitle & " " & "Version " & mstrVersion

            Me.CenterToScreen() 'Put this form in the center of the screen

            'Initialize the user input textboxes
            ZoTextBox.Text = CStr(InterFormMessage.dblZo)
            FreqTextBox.Text = CStr((InterFormMessage.dblFreq) / 1000000)

        Catch

            MessageBox.Show("Error in initialization", "Error", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click

        Try
            If ZoTextBox.Text <> "" Then
                InterFormMessage.dblZo = CDbl(ZoTextBox.Text)
            End If
        Catch
            MessageBox.Show("Please input a valid numeric value", "Error", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Try
            If FreqTextBox.Text <> "" Then
                InterFormMessage.dblFreq = CDbl(FreqTextBox.Text) * 1000000 'in MHz
            End If
        Catch
            MessageBox.Show("Please input a valid numeric value", "Error", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Close()

    End Sub
End Class
