Option Strict On

Public Class MainForm
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
    Friend WithEvents MainFormMenu As System.Windows.Forms.MainMenu
    Friend WithEvents OptionMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents AboutMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents FileMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents FileExitMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents OptionSettingMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents MainFormStatusBar As System.Windows.Forms.StatusBar
    Friend WithEvents MainFormStatusBarZo As System.Windows.Forms.StatusBarPanel
    Friend WithEvents MainFormStatusBarFrequency As System.Windows.Forms.StatusBarPanel
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents ActionMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents ActionSmithChartMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents ActionImpTranformMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents FilePrintMenuItem As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.MainFormMenu = New System.Windows.Forms.MainMenu(Me.components)
        Me.FileMenuItem = New System.Windows.Forms.MenuItem()
        Me.FilePrintMenuItem = New System.Windows.Forms.MenuItem()
        Me.FileExitMenuItem = New System.Windows.Forms.MenuItem()
        Me.ActionMenuItem = New System.Windows.Forms.MenuItem()
        Me.ActionSmithChartMenuItem = New System.Windows.Forms.MenuItem()
        Me.ActionImpTranformMenuItem = New System.Windows.Forms.MenuItem()
        Me.OptionMenuItem = New System.Windows.Forms.MenuItem()
        Me.OptionSettingMenuItem = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.AboutMenuItem = New System.Windows.Forms.MenuItem()
        Me.MainFormStatusBar = New System.Windows.Forms.StatusBar()
        Me.MainFormStatusBarZo = New System.Windows.Forms.StatusBarPanel()
        Me.MainFormStatusBarFrequency = New System.Windows.Forms.StatusBarPanel()
        CType(Me.MainFormStatusBarZo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainFormStatusBarFrequency, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainFormMenu
        '
        Me.MainFormMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.FileMenuItem, Me.ActionMenuItem, Me.OptionMenuItem, Me.MenuItem2, Me.AboutMenuItem})
        '
        'FileMenuItem
        '
        Me.FileMenuItem.Index = 0
        Me.FileMenuItem.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.FilePrintMenuItem, Me.FileExitMenuItem})
        Me.FileMenuItem.Text = "&File"
        '
        'FilePrintMenuItem
        '
        Me.FilePrintMenuItem.Index = 0
        Me.FilePrintMenuItem.Text = "&Print"
        '
        'FileExitMenuItem
        '
        Me.FileExitMenuItem.Index = 1
        Me.FileExitMenuItem.Text = "E&xit"
        '
        'ActionMenuItem
        '
        Me.ActionMenuItem.Index = 1
        Me.ActionMenuItem.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.ActionSmithChartMenuItem, Me.ActionImpTranformMenuItem})
        Me.ActionMenuItem.Text = "A&ction"
        '
        'ActionSmithChartMenuItem
        '
        Me.ActionSmithChartMenuItem.Index = 0
        Me.ActionSmithChartMenuItem.Text = "Show Smith Chart"
        '
        'ActionImpTranformMenuItem
        '
        Me.ActionImpTranformMenuItem.Index = 1
        Me.ActionImpTranformMenuItem.Text = "Perform Impedance Transform"
        '
        'OptionMenuItem
        '
        Me.OptionMenuItem.Index = 2
        Me.OptionMenuItem.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.OptionSettingMenuItem})
        Me.OptionMenuItem.Text = "&Options"
        '
        'OptionSettingMenuItem
        '
        Me.OptionSettingMenuItem.Index = 0
        Me.OptionSettingMenuItem.Text = "Se&ttings"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 3
        Me.MenuItem2.Text = "&Help"
        '
        'AboutMenuItem
        '
        Me.AboutMenuItem.Index = 4
        Me.AboutMenuItem.Text = "&About"
        '
        'MainFormStatusBar
        '
        Me.MainFormStatusBar.Location = New System.Drawing.Point(0, -24)
        Me.MainFormStatusBar.Name = "MainFormStatusBar"
        Me.MainFormStatusBar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.MainFormStatusBarZo, Me.MainFormStatusBarFrequency})
        Me.MainFormStatusBar.ShowPanels = True
        Me.MainFormStatusBar.Size = New System.Drawing.Size(352, 24)
        Me.MainFormStatusBar.TabIndex = 1
        Me.MainFormStatusBar.TabStop = True
        '
        'MainFormStatusBarZo
        '
        Me.MainFormStatusBarZo.Name = "MainFormStatusBarZo"
        '
        'MainFormStatusBarFrequency
        '
        Me.MainFormStatusBarFrequency.Name = "MainFormStatusBarFrequency"
        Me.MainFormStatusBarFrequency.Width = 180
        '
        'MainForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(352, 0)
        Me.Controls.Add(Me.MainFormStatusBar)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainFormMenu
        Me.Name = "MainForm"
        Me.Text = "Main Form"
        CType(Me.MainFormStatusBarZo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainFormStatusBarFrequency, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Author:            Fabian Kung Wai Lee
    'Last modified:     2nd March 2016
    'Description:       Main form codes for RF CAD software       
    'Filename:          MainForm.vb
    'Language:          Visual Basic .NET
    'Development tool:  Microsoft Visual Studio .NET 2013 Express

    '--- Module level declarations ---
    'Information about this software
    Private mstrFormTitle As String = "RF CAD"
    Private mstrVersion As String = "0.90"
    Private mstrAuthor As String = "Fabian Kung Wai Lee"
    Private mstrDate As String = "8th March 2016"

    'Variables
    Private mdblZo As Double = 50.0 'Reference impedance
    Private mdblfreq As Double = 430000000.0 'Default operating frequency in Hz

    'Objects (These are all user defined classes)
    Public InterFormMessage As New MessageClass
    Public SmithChartWindow As New SmithChartForm
    Private ImpedanceTransformWindow As New ImpedanceTransformForm

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Initialization routines here

        Try
            'Update the title and version of the main window
            Me.Text = mstrFormTitle & " " & "Version " & mstrVersion

            'Initialize inter-form communication class object
            'Variables:
            With InterFormMessage
                .intID = 1
                .dblZo = mdblZo
                .dblFreq = mdblfreq
            End With

            'Environment:
            With InterFormMessage
                .backgroundcolor = Me.BackColor
                .Linecolor1 = Color.Black
                .Linecolor2 = Color.Brown
                .intLineThickness = 2
                .Pointcolor = Color.Red
                .intPointWeigth = 6
            End With

            'Set the title and version of the SmithChart Window
            SmithChartWindow.mstrFormTitle = mstrFormTitle
            SmithChartWindow.mstrVersion = mstrVersion & " - Smith Chart"
            'Show Smith Chart window
            SmithChartWindow.Show()

        Catch

            MessageBox.Show("Error in initialization", "Error", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub AboutMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutMenuItem.Click

        Dim MainAboutForm As New AboutForm

        Try
            'Set the title and version of the About Window
            MainAboutForm.mstrFormTitle = mstrFormTitle
            MainAboutForm.mstrVersion = mstrVersion

            'Set the author and date of the About Window
            MainAboutForm.mstrAuthor = mstrAuthor
            MainAboutForm.mstrDate = mstrDate

            MainAboutForm.ShowDialog(Me) 'Display About Box in Modal mode.

        Catch

            MessageBox.Show("Error in displaying window", "Error", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub FileExitMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileExitMenuItem.Click

        Close()

    End Sub

    Private Sub ActionSmithChartMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ActionSmithChartMenuItem.Click

        Try
            'If the ID=0, this implies the Smith Chart Form
            'has been closed by the user.  So a new form has to be created.
            If InterFormMessage.intID = 0 Then
                SmithChartWindow = New SmithChartForm
                InterFormMessage.intID = 1
            End If

            'Set the title and version of the SmithChart Window
            SmithChartWindow.mstrFormTitle = mstrFormTitle
            SmithChartWindow.mstrVersion = mstrVersion & " - Smith Chart"
            'Show Smith Chart window
            SmithChartWindow.Show()
        Catch

            MessageBox.Show("Error in displaying window", "Error", _
                      MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub ActionImpTranformMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ActionImpTranformMenuItem.Click

        'Set the title and version of the Impedance Transform Window
        ImpedanceTransformWindow.mstrFormTitle = mstrFormTitle
        ImpedanceTransformWindow.mstrVersion = mstrVersion & " - Pi Network Transform"
        'Show Impedance Transform window
        ImpedanceTransformWindow.Show()

    End Sub

    Private Sub OptionSettingMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionSettingMenuItem.Click
        Dim MainOptionSettingForm As New OptionForm

        Try
            'Set the title and version of the Setting Window
            MainOptionSettingForm.mstrFormTitle = mstrFormTitle
            MainOptionSettingForm.mstrVersion = mstrVersion
            MainOptionSettingForm.ShowDialog(Me) 'Show form in Modal mode.
            'Update module level variables and status bar display whenever focus is
            'returned to Main Form
            mdblZo = InterFormMessage.dblZo
            SmithChartWindow.MainStatusBarZo.Text = "Zo = " & CStr(mdblZo) & " Ohm"
            mdblfreq = InterFormMessage.dblFreq
            SmithChartWindow.MainStatusBarFreq.Text = "Center frequency = " & CStr(mdblfreq / 1000000) & " MHz"

            ImpedanceTransformWindow.FrequencyLabel.Text = CStr(mdblfreq / 1000000) & " MHz"
            ImpedanceTransformWindow.ZoLabel.Text = CStr(mdblZo) & " Ohm"
        Catch

            MessageBox.Show("Error in displaying window", "Error", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub



End Class

