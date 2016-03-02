Public Class AboutForm
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
    Friend WithEvents CloseButton As System.Windows.Forms.Button
    Friend WithEvents AuthorLabel As System.Windows.Forms.Label
    Friend WithEvents DateLabel As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LicenesLabel As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(AboutForm))
        Me.CloseButton = New System.Windows.Forms.Button
        Me.AuthorLabel = New System.Windows.Forms.Label
        Me.DateLabel = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.LicenesLabel = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'CloseButton
        '
        Me.CloseButton.Location = New System.Drawing.Point(72, 232)
        Me.CloseButton.Name = "CloseButton"
        Me.CloseButton.Size = New System.Drawing.Size(104, 40)
        Me.CloseButton.TabIndex = 0
        Me.CloseButton.Text = "&Agree"
        '
        'AuthorLabel
        '
        Me.AuthorLabel.Location = New System.Drawing.Point(80, 32)
        Me.AuthorLabel.Name = "AuthorLabel"
        Me.AuthorLabel.Size = New System.Drawing.Size(176, 24)
        Me.AuthorLabel.TabIndex = 1
        '
        'DateLabel
        '
        Me.DateLabel.Location = New System.Drawing.Point(80, 64)
        Me.DateLabel.Name = "DateLabel"
        Me.DateLabel.Size = New System.Drawing.Size(136, 16)
        Me.DateLabel.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Written by"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Date"
        '
        'LicenesLabel
        '
        Me.LicenesLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LicenesLabel.Location = New System.Drawing.Point(24, 120)
        Me.LicenesLabel.Name = "LicenesLabel"
        Me.LicenesLabel.Size = New System.Drawing.Size(216, 96)
        Me.LicenesLabel.TabIndex = 5
        Me.LicenesLabel.Text = "This is a freeware.  As such you are free to distribute this software in any mann" & _
        "er you wish.  The author does not assume any responsibility for losses, damages " & _
        "and injuries as a result of using this software."
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 24)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "License:"
        '
        'AboutForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(264, 286)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.LicenesLabel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DateLabel)
        Me.Controls.Add(Me.AuthorLabel)
        Me.Controls.Add(Me.CloseButton)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "AboutForm"
        Me.Text = "AboutForm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Module level declarations
    Public mstrFormTitle As String
    Public mstrVersion As String
    Public mstrAuthor As String
    Public mstrDate As String

    Private Sub AboutForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            'Update the title and version of the main window
            Me.Text = "About " & mstrFormTitle & " " & "Version " & mstrVersion

            'Tradisionally About Box cannot be resize, minimize and maximize.  Also an
            'About Box is usually modal in display mode.
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MinimizeBox = False
            Me.MaximizeBox = False

            'Update author and date information
            Me.AuthorLabel.Text = mstrAuthor
            Me.DateLabel.Text = mstrDate

            Me.CenterToScreen() 'Put this form in the center of the screen
        Catch

            MessageBox.Show("Error in initialization", "Error", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub CloseButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseButton.Click

        Close()

    End Sub

End Class
