Imports System.Math

Public Class SmithChartForm
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
    Friend WithEvents MainPicBox As System.Windows.Forms.PictureBox
    Friend WithEvents RLabel As System.Windows.Forms.Label
    Friend WithEvents XLabel As System.Windows.Forms.Label
    Friend WithEvents GLabel As System.Windows.Forms.Label
    Friend WithEvents BLabel As System.Windows.Forms.Label
    Friend WithEvents TLabel As System.Windows.Forms.Label
    Friend WithEvents PLabel As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SmithHScrollBar As System.Windows.Forms.HScrollBar
    Friend WithEvents SmithVScrollBar As System.Windows.Forms.VScrollBar
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents PlotButton As System.Windows.Forms.Button
    Friend WithEvents ClearButton As System.Windows.Forms.Button
    Friend WithEvents CoordinateGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents FitButton As System.Windows.Forms.Button
    Friend WithEvents ZCoorCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents YCoorCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents MainToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents InputModeComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents Input1Label As System.Windows.Forms.Label
    Friend WithEvents Input2Label As System.Windows.Forms.Label
    Friend WithEvents Input1TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Input2TextBox As System.Windows.Forms.TextBox
    Friend WithEvents MainStatusBar As System.Windows.Forms.StatusBar
    Friend WithEvents MainStatusBarZo As System.Windows.Forms.StatusBarPanel
    Friend WithEvents MainStatusBarFreq As System.Windows.Forms.StatusBarPanel
    Friend WithEvents SmithChartContextMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents RCircleMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents XCircleMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents GCircleMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents Bar1MenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents BCircleMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents Bar2MenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents PlotMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents ClearMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents Bar3MenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents Bar4MenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents DynamicYCoorMenuItem As System.Windows.Forms.MenuItem
    Friend WithEvents DynamicZCoorMenuItem As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SmithChartForm))
        Me.MainPicBox = New System.Windows.Forms.PictureBox
        Me.SmithChartContextMenu = New System.Windows.Forms.ContextMenu
        Me.RCircleMenuItem = New System.Windows.Forms.MenuItem
        Me.GCircleMenuItem = New System.Windows.Forms.MenuItem
        Me.Bar1MenuItem = New System.Windows.Forms.MenuItem
        Me.XCircleMenuItem = New System.Windows.Forms.MenuItem
        Me.BCircleMenuItem = New System.Windows.Forms.MenuItem
        Me.Bar2MenuItem = New System.Windows.Forms.MenuItem
        Me.PlotMenuItem = New System.Windows.Forms.MenuItem
        Me.Bar3MenuItem = New System.Windows.Forms.MenuItem
        Me.ClearMenuItem = New System.Windows.Forms.MenuItem
        Me.Bar4MenuItem = New System.Windows.Forms.MenuItem
        Me.DynamicZCoorMenuItem = New System.Windows.Forms.MenuItem
        Me.DynamicYCoorMenuItem = New System.Windows.Forms.MenuItem
        Me.RLabel = New System.Windows.Forms.Label
        Me.XLabel = New System.Windows.Forms.Label
        Me.GLabel = New System.Windows.Forms.Label
        Me.BLabel = New System.Windows.Forms.Label
        Me.TLabel = New System.Windows.Forms.Label
        Me.PLabel = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.SmithHScrollBar = New System.Windows.Forms.HScrollBar
        Me.SmithVScrollBar = New System.Windows.Forms.VScrollBar
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.PlotButton = New System.Windows.Forms.Button
        Me.Input1TextBox = New System.Windows.Forms.TextBox
        Me.Input2TextBox = New System.Windows.Forms.TextBox
        Me.ClearButton = New System.Windows.Forms.Button
        Me.CoordinateGroupBox = New System.Windows.Forms.GroupBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.YCoorCheckBox = New System.Windows.Forms.CheckBox
        Me.ZCoorCheckBox = New System.Windows.Forms.CheckBox
        Me.FitButton = New System.Windows.Forms.Button
        Me.MainToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.InputModeComboBox = New System.Windows.Forms.ComboBox
        Me.Input1Label = New System.Windows.Forms.Label
        Me.Input2Label = New System.Windows.Forms.Label
        Me.MainStatusBar = New System.Windows.Forms.StatusBar
        Me.MainStatusBarZo = New System.Windows.Forms.StatusBarPanel
        Me.MainStatusBarFreq = New System.Windows.Forms.StatusBarPanel
        Me.CoordinateGroupBox.SuspendLayout()
        CType(Me.MainStatusBarZo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainStatusBarFreq, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainPicBox
        '
        Me.MainPicBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MainPicBox.ContextMenu = Me.SmithChartContextMenu
        Me.MainPicBox.Location = New System.Drawing.Point(80, 24)
        Me.MainPicBox.Name = "MainPicBox"
        Me.MainPicBox.Size = New System.Drawing.Size(424, 384)
        Me.MainPicBox.TabIndex = 0
        Me.MainPicBox.TabStop = False
        '
        'SmithChartContextMenu
        '
        Me.SmithChartContextMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.RCircleMenuItem, Me.GCircleMenuItem, Me.Bar1MenuItem, Me.XCircleMenuItem, Me.BCircleMenuItem, Me.Bar2MenuItem, Me.PlotMenuItem, Me.Bar3MenuItem, Me.ClearMenuItem, Me.Bar4MenuItem, Me.DynamicZCoorMenuItem, Me.DynamicYCoorMenuItem})
        '
        'RCircleMenuItem
        '
        Me.RCircleMenuItem.Index = 0
        Me.RCircleMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlR
        Me.RCircleMenuItem.Text = "Insert Extra R Circle"
        '
        'GCircleMenuItem
        '
        Me.GCircleMenuItem.Index = 1
        Me.GCircleMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlG
        Me.GCircleMenuItem.Text = "Insert Extra G Circle"
        '
        'Bar1MenuItem
        '
        Me.Bar1MenuItem.Index = 2
        Me.Bar1MenuItem.Text = "-"
        '
        'XCircleMenuItem
        '
        Me.XCircleMenuItem.Index = 3
        Me.XCircleMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlX
        Me.XCircleMenuItem.Text = "Insert Extra X Circle"
        '
        'BCircleMenuItem
        '
        Me.BCircleMenuItem.Index = 4
        Me.BCircleMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlB
        Me.BCircleMenuItem.Text = "Insert Extra B Circle"
        '
        'Bar2MenuItem
        '
        Me.Bar2MenuItem.Index = 5
        Me.Bar2MenuItem.Text = "-"
        '
        'PlotMenuItem
        '
        Me.PlotMenuItem.Index = 6
        Me.PlotMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlP
        Me.PlotMenuItem.Text = "Plot Reflection Coefficient"
        '
        'Bar3MenuItem
        '
        Me.Bar3MenuItem.Index = 7
        Me.Bar3MenuItem.Text = "-"
        '
        'ClearMenuItem
        '
        Me.ClearMenuItem.Index = 8
        Me.ClearMenuItem.Shortcut = System.Windows.Forms.Shortcut.Del
        Me.ClearMenuItem.Text = "Clear Extra Circles"
        '
        'Bar4MenuItem
        '
        Me.Bar4MenuItem.Index = 9
        Me.Bar4MenuItem.Text = "-"
        '
        'DynamicZCoorMenuItem
        '
        Me.DynamicZCoorMenuItem.Index = 10
        Me.DynamicZCoorMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlZ
        Me.DynamicZCoorMenuItem.Text = "Enable Dynamic Z Coordinate"
        '
        'DynamicYCoorMenuItem
        '
        Me.DynamicYCoorMenuItem.Index = 11
        Me.DynamicYCoorMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlY
        Me.DynamicYCoorMenuItem.Text = "Enable Dynamic Y Coordinate"
        '
        'RLabel
        '
        Me.RLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.RLabel.Location = New System.Drawing.Point(24, 184)
        Me.RLabel.Name = "RLabel"
        Me.RLabel.Size = New System.Drawing.Size(48, 16)
        Me.RLabel.TabIndex = 1
        '
        'XLabel
        '
        Me.XLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.XLabel.Location = New System.Drawing.Point(24, 208)
        Me.XLabel.Name = "XLabel"
        Me.XLabel.Size = New System.Drawing.Size(48, 16)
        Me.XLabel.TabIndex = 2
        '
        'GLabel
        '
        Me.GLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.GLabel.Location = New System.Drawing.Point(24, 232)
        Me.GLabel.Name = "GLabel"
        Me.GLabel.Size = New System.Drawing.Size(48, 16)
        Me.GLabel.TabIndex = 3
        '
        'BLabel
        '
        Me.BLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.BLabel.Location = New System.Drawing.Point(24, 256)
        Me.BLabel.Name = "BLabel"
        Me.BLabel.Size = New System.Drawing.Size(48, 16)
        Me.BLabel.TabIndex = 4
        '
        'TLabel
        '
        Me.TLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TLabel.Location = New System.Drawing.Point(24, 280)
        Me.TLabel.Name = "TLabel"
        Me.TLabel.Size = New System.Drawing.Size(48, 16)
        Me.TLabel.TabIndex = 5
        '
        'PLabel
        '
        Me.PLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PLabel.Location = New System.Drawing.Point(24, 304)
        Me.PLabel.Name = "PLabel"
        Me.PLabel.Size = New System.Drawing.Size(48, 16)
        Me.PLabel.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 280)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 16)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "|T|"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(0, 304)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 24)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "<T"
        '
        'SmithHScrollBar
        '
        Me.SmithHScrollBar.Location = New System.Drawing.Point(88, 424)
        Me.SmithHScrollBar.Name = "SmithHScrollBar"
        Me.SmithHScrollBar.Size = New System.Drawing.Size(440, 16)
        Me.SmithHScrollBar.TabIndex = 12
        '
        'SmithVScrollBar
        '
        Me.SmithVScrollBar.Location = New System.Drawing.Point(512, 24)
        Me.SmithVScrollBar.Name = "SmithVScrollBar"
        Me.SmithVScrollBar.Size = New System.Drawing.Size(16, 392)
        Me.SmithVScrollBar.TabIndex = 13
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(0, 232)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(16, 16)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "G"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(0, 256)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 16)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "B"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(0, 184)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(24, 24)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "R"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(0, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(16, 16)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "X"
        '
        'PlotButton
        '
        Me.PlotButton.Location = New System.Drawing.Point(8, 104)
        Me.PlotButton.Name = "PlotButton"
        Me.PlotButton.Size = New System.Drawing.Size(48, 32)
        Me.PlotButton.TabIndex = 19
        Me.PlotButton.Text = "&Plot"
        Me.MainToolTip.SetToolTip(Me.PlotButton, "Plot the Z or Y or Rho on the Smith Chart")
        '
        'Input1TextBox
        '
        Me.Input1TextBox.Location = New System.Drawing.Point(32, 40)
        Me.Input1TextBox.Name = "Input1TextBox"
        Me.Input1TextBox.Size = New System.Drawing.Size(40, 20)
        Me.Input1TextBox.TabIndex = 20
        Me.Input1TextBox.Text = ""
        Me.MainToolTip.SetToolTip(Me.Input1TextBox, "Key in numeric value")
        '
        'Input2TextBox
        '
        Me.Input2TextBox.Location = New System.Drawing.Point(32, 72)
        Me.Input2TextBox.Name = "Input2TextBox"
        Me.Input2TextBox.Size = New System.Drawing.Size(40, 20)
        Me.Input2TextBox.TabIndex = 21
        Me.Input2TextBox.Text = ""
        Me.MainToolTip.SetToolTip(Me.Input2TextBox, "Key in numeric value")
        '
        'ClearButton
        '
        Me.ClearButton.Location = New System.Drawing.Point(8, 136)
        Me.ClearButton.Name = "ClearButton"
        Me.ClearButton.Size = New System.Drawing.Size(48, 32)
        Me.ClearButton.TabIndex = 22
        Me.ClearButton.Text = "&Clear Plot"
        Me.MainToolTip.SetToolTip(Me.ClearButton, "Clear all points on the Smith Chart")
        '
        'CoordinateGroupBox
        '
        Me.CoordinateGroupBox.Controls.Add(Me.Label8)
        Me.CoordinateGroupBox.Controls.Add(Me.Label7)
        Me.CoordinateGroupBox.Controls.Add(Me.YCoorCheckBox)
        Me.CoordinateGroupBox.Controls.Add(Me.ZCoorCheckBox)
        Me.CoordinateGroupBox.Controls.Add(Me.FitButton)
        Me.CoordinateGroupBox.Location = New System.Drawing.Point(0, 328)
        Me.CoordinateGroupBox.Name = "CoordinateGroupBox"
        Me.CoordinateGroupBox.Size = New System.Drawing.Size(80, 120)
        Me.CoordinateGroupBox.TabIndex = 27
        Me.CoordinateGroupBox.TabStop = False
        Me.CoordinateGroupBox.Text = "Coordinate"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(32, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(24, 24)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "Y"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(32, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(24, 24)
        Me.Label7.TabIndex = 33
        Me.Label7.Text = "Z"
        '
        'YCoorCheckBox
        '
        Me.YCoorCheckBox.Location = New System.Drawing.Point(8, 48)
        Me.YCoorCheckBox.Name = "YCoorCheckBox"
        Me.YCoorCheckBox.Size = New System.Drawing.Size(16, 16)
        Me.YCoorCheckBox.TabIndex = 32
        Me.MainToolTip.SetToolTip(Me.YCoorCheckBox, "Show Y coordinate")
        '
        'ZCoorCheckBox
        '
        Me.ZCoorCheckBox.Checked = True
        Me.ZCoorCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ZCoorCheckBox.Location = New System.Drawing.Point(8, 24)
        Me.ZCoorCheckBox.Name = "ZCoorCheckBox"
        Me.ZCoorCheckBox.Size = New System.Drawing.Size(16, 16)
        Me.ZCoorCheckBox.TabIndex = 31
        Me.MainToolTip.SetToolTip(Me.ZCoorCheckBox, "Show Z coordinate")
        '
        'FitButton
        '
        Me.FitButton.Location = New System.Drawing.Point(8, 72)
        Me.FitButton.Name = "FitButton"
        Me.FitButton.Size = New System.Drawing.Size(56, 32)
        Me.FitButton.TabIndex = 30
        Me.FitButton.Text = "&Fit Window"
        Me.MainToolTip.SetToolTip(Me.FitButton, "Fit the Smith Chart into the window")
        '
        'InputModeComboBox
        '
        Me.InputModeComboBox.Items.AddRange(New Object() {"Z", "Y", "Rho"})
        Me.InputModeComboBox.Location = New System.Drawing.Point(8, 8)
        Me.InputModeComboBox.Name = "InputModeComboBox"
        Me.InputModeComboBox.Size = New System.Drawing.Size(64, 21)
        Me.InputModeComboBox.TabIndex = 28
        Me.MainToolTip.SetToolTip(Me.InputModeComboBox, "Select input mode")
        '
        'Input1Label
        '
        Me.Input1Label.Location = New System.Drawing.Point(8, 40)
        Me.Input1Label.Name = "Input1Label"
        Me.Input1Label.Size = New System.Drawing.Size(24, 24)
        Me.Input1Label.TabIndex = 29
        '
        'Input2Label
        '
        Me.Input2Label.Location = New System.Drawing.Point(8, 72)
        Me.Input2Label.Name = "Input2Label"
        Me.Input2Label.Size = New System.Drawing.Size(24, 24)
        Me.Input2Label.TabIndex = 30
        '
        'MainStatusBar
        '
        Me.MainStatusBar.Location = New System.Drawing.Point(0, 466)
        Me.MainStatusBar.Name = "MainStatusBar"
        Me.MainStatusBar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.MainStatusBarZo, Me.MainStatusBarFreq})
        Me.MainStatusBar.ShowPanels = True
        Me.MainStatusBar.Size = New System.Drawing.Size(544, 20)
        Me.MainStatusBar.TabIndex = 31
        '
        'MainStatusBarFreq
        '
        Me.MainStatusBarFreq.Width = 160
        '
        'SmithChartForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(544, 486)
        Me.Controls.Add(Me.MainStatusBar)
        Me.Controls.Add(Me.Input2Label)
        Me.Controls.Add(Me.Input1Label)
        Me.Controls.Add(Me.InputModeComboBox)
        Me.Controls.Add(Me.CoordinateGroupBox)
        Me.Controls.Add(Me.ClearButton)
        Me.Controls.Add(Me.Input2TextBox)
        Me.Controls.Add(Me.Input1TextBox)
        Me.Controls.Add(Me.PlotButton)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.SmithVScrollBar)
        Me.Controls.Add(Me.SmithHScrollBar)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PLabel)
        Me.Controls.Add(Me.TLabel)
        Me.Controls.Add(Me.BLabel)
        Me.Controls.Add(Me.GLabel)
        Me.Controls.Add(Me.XLabel)
        Me.Controls.Add(Me.RLabel)
        Me.Controls.Add(Me.MainPicBox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(200, 200)
        Me.Name = "SmithChartForm"
        Me.CoordinateGroupBox.ResumeLayout(False)
        CType(Me.MainStatusBarZo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainStatusBarFreq, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Author:            Fabian Kung Wai Lee
    'Last modified:     1st March 2016
    'Description:       Codes to generate a Smith Chart window and other functions     
    'Filename:          SmithChartForm.vb
    'Language:          Visual Basic .NET
    'Development tool:  Microsoft Visual Studio .NET Express 2013

    '--- Module level declarations ---
    'Environment:
    Public mstrFormTitle As String
    Public mstrVersion As String
    Public mobjInterFormMessage As New MessageClass

    'Impedance, admittance and reflection coefficient that other modules 
    'can access to:
    'Note: This is the Z, Y or T associated with the mouse cursor on the Smith Chart
    'window that other modules can access to.
    Public mobjT As New ComplexNumberClass(0, 0) 'Reflection coefficient
    Public mobjZ As New ComplexNumberClass(0, 0) 'Impedance
    Public mobjY As New ComplexNumberClass(0, 0) 'Admittance

    'Note: Auxiliary reflection coefficient which we can plot on the Smith Chart.  An external
    'module can perform some analysis and display their results in the Smith Chart by manipulating
    'this public object (of course one must set the property 'visible' to true to display it)
    Public mobjTaux As New ComplexNumberClass(0, 0)

    'Reflection coefficients objects
    Private mobj1 As New ComplexNumberClass(1, 0) 'constant: 1+j0
    Private mobjTo(20) As ComplexNumberClass 'Internal reflection coefficients array.
    Private mintCount As Integer = 0    'Internal reflection coefficient array pointer.
    Const MAXCOUNT As Integer = 18   'The number of internal reflection coefficients.

    'Smith chart coordinate and related properties, in pixels.
    Private mintOXw As Integer  'Real axis origin in Windows coordinate system
    Private mintOYw As Integer 'Imaginary axis origin in Windows coordinate system
    Private mintOXwOffset As Integer 'Zero offset for real axis origin
    Private mintOYwOffset As Integer 'Zero offset for imaginary axis origin

    Private mintPicBoxStartX As Integer '(x,y) coordinate for the top left hand corner
    Private mintPicBoxStartY As Integer 'of the main picturebox, i.e. the start point.

    'Size of the Smith Chart boundary, this is defined as half of the height or
    'width of the graphics window, whichever is larger, it is specified in pixels.
    Private mintSpan As Integer 'Radius or size of the Smith Chart in pixels.
    Private mintDeltaSpan As Integer 'Change in the size of the Smith Chart in pixels.

    Private mLinePen1 As Pen 'Pen to draw impedance coordinate.
    Private mLinePen2 As Pen 'Pen to draw admittance coordinate.
    Private mLinePen3 As Pen 'Pen to draw miscellaneous items.
    Private mPointbrush As SolidBrush 'Brush to draw points.
    Private mintInputMode As Integer '0 = Z input, 1 = Y input, others = Rho input.

    'Miscellaneous module/class level variables
    Private mintHSBarOldValue As Integer
    Private mintVSBarOldValue As Integer

    Private mlstRCircleList As New ArrayList 'Array List for extra Constant R Circles.
    Private mlstXCircleList As New ArrayList 'Array List for extra Constant X Circles.
    Private mlstGCircleList As New ArrayList 'Array List for extra Constant G Circles.
    Private mlstBCircleList As New ArrayList 'Array List for extra Constant B Circles.

    Private Sub SmithChartForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim picboxstartpoint As Point
        Dim HSbarstartpoint As Point
        Dim I As Integer

        Try
            'Update the title and version of the main window
            Me.Text = mstrFormTitle & " " & "Version " & mstrVersion

            'Initialize Pen and Brush for 2D graphics
            mLinePen1 = New Pen(mobjInterFormMessage.Linecolor1, mobjInterFormMessage.intLineThickness)
            mLinePen2 = New Pen(mobjInterFormMessage.Linecolor2, mobjInterFormMessage.intLineThickness)
            mLinePen3 = New Pen(Color.Azure, mobjInterFormMessage.intLineThickness)
            mPointbrush = New SolidBrush(mobjInterFormMessage.Pointcolor)

            'Set the main picture box size and location
            mintPicBoxStartX = 80
            mintPicBoxStartY = 15
            picboxstartpoint = New Point(mintPicBoxStartX, mintPicBoxStartY)
            MainPicBox.Width = Me.Width - 118
            MainPicBox.Height = Me.Height - 95
            MainPicBox.Location = picboxstartpoint

            'Determine the size of the Smith Chart and its origin with respect
            'to the picture box
            mintDeltaSpan = 0
            'Finding the radius:
            If MainPicBox.Width > MainPicBox.Height Then
                mintSpan = (MainPicBox.Height / 2) - 4 'allow some clearance from the side
            Else
                mintSpan = (MainPicBox.Width / 2) - 4 'allow some clearance from the side
            End If

            'Setting the origin of the main picturebox coordinate system
            mintOXw = MainPicBox.Width / 2
            mintOYw = MainPicBox.Height / 2
            mintOXwOffset = 0
            mintOYwOffset = 0

            'Set the horizontal scrollbar size, location and parameters
            SmithHScrollBar.Left = picboxstartpoint.X
            SmithHScrollBar.Top = picboxstartpoint.Y + MainPicBox.Height + 5
            SmithHScrollBar.Width = MainPicBox.Width
            SmithHScrollBar.Height = 16
            SmithHScrollBar.Maximum = MainPicBox.Width
            SmithHScrollBar.Value = MainPicBox.Width / 2
            mintHSBarOldValue = SmithHScrollBar.Value

            'Set the vertical scrollbar size and location
            SmithVScrollBar.Left = picboxstartpoint.X + MainPicBox.Width + 5
            SmithVScrollBar.Top = picboxstartpoint.Y
            SmithVScrollBar.Width = 16
            SmithVScrollBar.Height = MainPicBox.Height
            SmithVScrollBar.Maximum = MainPicBox.Height
            SmithVScrollBar.Value = MainPicBox.Height / 2
            mintVSBarOldValue = SmithVScrollBar.Value

            'Initialize other miscellaneous objects
            For I = 0 To MAXCOUNT
                'Instantiate each object inside the array
                mobjTo(I) = New ComplexNumberClass(0, 0)
            Next
            InputModeComboBox.SelectedIndex = 0
            mintInputMode = 0 'Z input mode.
            'Update Status Bar.
            MainStatusBarZo.Text = "Zo = " & CStr(mobjInterformmessage.dblZo) & " Ohm"
            MainStatusBarFreq.Text = "Center frequency = " & _
            CStr(mobjInterFormMessage.dblFreq / 1000000) & " MHz"

        Catch theException As Exception

            MessageBox.Show("SmithChartForm_Load " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub SmithChartForm_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Dim picboxstartpoint As New Point(80, 15)

        Try
            'Set the minimum width and height of the form
            'If it is too small the other routines to plot Smith Chart lines
            'will produce error
            If Me.Width < 240 Then
                Me.Width = 240
            End If
            If Me.Height < 200 Then
                Me.Height = 200
            End If

            'Set the main picture box size and location
            MainPicBox.Width = Me.Width - 118
            MainPicBox.Height = Me.Height - 95
            MainPicBox.Location = picboxstartpoint

            'Determine the size of the Smith Chart and its origin with respect
            'to the picture box
            'Finding the radius:
            'If MainPicBox.Width > MainPicBox.Height Then
            'mintSpan = (MainPicBox.Height / 2) - 4 'allow some clearance from the side
            'Else
            '   mintSpan = (MainPicBox.Width / 2) - 4 'allow some clearance from the side
            'End If

            'mintSpan = mintSpan + mintDeltaSpan 'Adding zoom factor to the radius

            'Setting the Origin of the main picturebox coordinate system
            mintOXw = MainPicBox.Width / 2
            mintOYw = MainPicBox.Height / 2

            'Set the horizontal scrollbar size, location and parameters
            SmithHScrollBar.Left = picboxstartpoint.X
            SmithHScrollBar.Top = picboxstartpoint.Y + MainPicBox.Height + 5
            SmithHScrollBar.Width = MainPicBox.Width
            SmithHScrollBar.Height = 16
            SmithHScrollBar.Maximum = MainPicBox.Width
            mintHSBarOldValue = SmithHScrollBar.Value

            'Set the vertical scrollbar size, location and parameters
            SmithVScrollBar.Left = picboxstartpoint.X + MainPicBox.Width + 5
            SmithVScrollBar.Top = picboxstartpoint.Y
            SmithVScrollBar.Width = 16
            SmithVScrollBar.Height = MainPicBox.Height
            SmithVScrollBar.Maximum = MainPicBox.Height
            mintVSBarOldValue = SmithVScrollBar.Value

            MainPicBox.Refresh() 'Regenerate the 2D graphics in the picture box.

        Catch theException As Exception

            MessageBox.Show("SmithChartForm_Resize " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub MainPicBox_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles MainPicBox.DoubleClick

        Dim MainDisplaySettingForm As New DisplaySettingForm

        Try
            'Set the title and version of the Display Setting Window
            MainDisplaySettingForm.mstrFormTitle = mstrFormTitle
            MainDisplaySettingForm.mstrVersion = mstrVersion
            MainDisplaySettingForm.ShowDialog(Me) 'Show form in modal mode. The user needs to close the form before further
            'execution of subsequent codes can proceed.

            'Update environment variables and picture box
            MainPicBox.BackColor = mobjInterFormMessage.backgroundcolor
            mLinePen1.Color = mobjInterFormMessage.Linecolor1
            mLinePen2.Color = mobjInterFormMessage.Linecolor2
            mLinePen1.Width = mobjInterFormMessage.intLineThickness
            mLinePen2.Width = mobjInterFormMessage.intLineThickness
            mPointbrush.Color = mobjInterFormMessage.Pointcolor

            MainPicBox.Refresh() 'Redraw picture box content

        Catch theException As Exception

            MessageBox.Show("MainPicBox_DoubleClick " & theException.Message, "Error", _
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    'Routines to perform when the mouse move within the main picture box, which display the Smith Chart
    Private Sub MainPicBox_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MainPicBox.MouseMove
        Try
            Dim objZ As New ComplexNumberClass(0, 0)
            Dim objY As New ComplexNumberClass(0, 0)
         
            'Convert mouse's coordinate in the picture box into reflection coefficient
            ConvertWinCoordinateToReflectionCoeff(e.X, e.Y)

            'Display reflection coefficient on window's labels
            TLabel.Text = CStr(FormatNumber(mobjT.Magnitude, 3)) '3 decimal places
            PLabel.Text = CStr(FormatNumber(mobjT.Phase, 3)) '3 decimal places

            'Display impedance on window's labels
            'Check for overflow (infinity) first.
            If ((mobjT.Real = 1) And (mobjT.Imagninary = 0)) Then
                ' Do nothing
            Else
                mobjZ = TtoZ(mobjT)
                RLabel.Text = CStr(FormatNumber(mobjZ.Real, 2)) '2 decimal places
                XLabel.Text = CStr(FormatNumber(mobjZ.Imagninary, 2)) '2 decimal places
            End If

            'Dispaly admittance on window's labels
            'Check for overflow (infinity) first.
            If ((mobjT.Real = -1) And (mobjT.Imagninary = 0)) Then
                ' Do nothing
            Else
                mobjY = mobjZ.DivComplex(mobj1, mobjZ) 'Y=1/Z
                GLabel.Text = CStr(FormatNumber(mobjY.Real, 5))  '5 decimal places
                BLabel.Text = CStr(FormatNumber(mobjY.Imagninary, 5)) '5 decimal places
            End If

            'Only refresh MainPicBox if dynamic coordinate is enabled.  This is to
            'prevent unnecessary computations.
            If (DynamicZCoorMenuItem.Checked = True) Or _
             (DynamicYCoorMenuItem.Checked = True) Then
                MainPicBox.Refresh()    'Redraw MainPicBox.
            End If

        Catch theException As Exception

            MessageBox.Show("MainPicBox_MouseMove " & theException.Message, "Error", _
                      MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MainPicBox_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MainPicBox.Paint
        Dim Grobj As Graphics = e.Graphics
        Dim intR As Integer = mintSpan
        Dim dblZo As Double = mobjInterFormMessage.dblZo
        Dim intPW As Integer = mobjInterFormMessage.intPointWeigth
        Dim pt As Point
        Dim I As Integer

        Try
            'Draw the boundary of the Smith Chart
            PlotConstRCircle(0.0, dblZo, Grobj)

            'Check if Show Impedance (Z) Coordinate option is selected
            If ZCoorCheckBox.Checked = True Then
                'Plot constant R circles
                PlotConstRCircle(20.0, dblZo, Grobj)
                PlotConstRCircle(50.0, dblZo, Grobj)
                PlotConstRCircle(100.0, dblZo, Grobj)
                PlotConstRCircle(200.0, dblZo, Grobj)
                PlotConstRCircle(400.0, dblZo, Grobj)
                'Plot constant X circles
                PlotConstXCircle(10.0, dblZo, Grobj)
                PlotConstXCircle(30.0, dblZo, Grobj)
                PlotConstXCircle(60.0, dblZo, Grobj)
                PlotConstXCircle(100.0, dblZo, Grobj)
                PlotConstXCircle(200.0, dblZo, Grobj)
                PlotConstXCircle(400.0, dblZo, Grobj)
                PlotConstXCircle(-10.0, dblZo, Grobj)
                PlotConstXCircle(-30.0, dblZo, Grobj)
                PlotConstXCircle(-60.0, dblZo, Grobj)
                PlotConstXCircle(-100.0, dblZo, Grobj)
                PlotConstXCircle(-200.0, dblZo, Grobj)
                PlotConstXCircle(-400.0, dblZo, Grobj)

            End If

            If YCoorCheckBox.Checked = True Then
                'Plot constant B circles
                PlotConstBCircle(0.004, dblZo, Grobj)
                PlotConstBCircle(0.012, dblZo, Grobj)
                PlotConstBCircle(0.024, dblZo, Grobj)
                PlotConstBCircle(0.04, dblZo, Grobj)
                PlotConstBCircle(0.08, dblZo, Grobj)
                PlotConstBCircle(0.16, dblZo, Grobj)
                PlotConstBCircle(-0.004, dblZo, Grobj)
                PlotConstBCircle(-0.012, dblZo, Grobj)
                PlotConstBCircle(-0.024, dblZo, Grobj)
                PlotConstBCircle(-0.04, dblZo, Grobj)
                PlotConstBCircle(-0.08, dblZo, Grobj)
                PlotConstBCircle(-0.16, dblZo, Grobj)

                'Plot constant G circles
                PlotConstGCircle(0.16, dblZo, Grobj)
                PlotConstGCircle(0.08, dblZo, Grobj)
                PlotConstGCircle(0.04, dblZo, Grobj)
                PlotConstGCircle(0.02, dblZo, Grobj)
                PlotConstGCircle(0.008, dblZo, Grobj)
            End If

            'Draw a horizontal line across Smith Chart
            If ZCoorCheckBox.Checked = True Then
                Grobj.DrawLine(mLinePen1, (mintOXw + mintOXwOffset) - intR, (mintOYw + mintOYwOffset), _
                (mintOXw + mintOXwOffset) + intR, (mintOYw + mintOYwOffset))
            ElseIf YCoorCheckBox.Checked = True Then
                Grobj.DrawLine(mLinePen2, (mintOXw + mintOXwOffset) - intR, (mintOYw + mintOYwOffset), _
                (mintOXw + mintOXwOffset) + intR, (mintOYw + mintOYwOffset))
            End If

            'Plot all complex reflection coefficient (if their property blnVisible
            'as asserted true)
            For I = 0 To MAXCOUNT
                If (mobjTo(I).Visible = True) Then
                    pt = ConvertReflectionCoeffToWinCoordinate(mobjTo(I))
                    pt.X = pt.X - (intPW / 2) 'To adjust the center position when point weigth is >1.
                    pt.Y = pt.Y - (intPW / 2)
                    Grobj.FillEllipse(mPointbrush, pt.X, pt.Y, mobjInterFormMessage.intPointWeigth, mobjInterFormMessage.intPointWeigth)
                End If
            Next

            'Plot auxiliary reflection coefficient
            If (mobjTaux.Visible = True) Then
                pt = ConvertReflectionCoeffToWinCoordinate(mobjTaux)
                pt.X = pt.X - (intPW / 2) 'To adjust the center position when point weigth is >1.
                pt.Y = pt.Y - (intPW / 2)
                Grobj.FillEllipse(mPointbrush, pt.X, pt.Y, mobjInterFormMessage.intPointWeigth, mobjInterFormMessage.intPointWeigth)
            End If

            'Plot all extra Constant R Circle in the mlstRCircleList Array List.
            For I = 0 To (mlstRCircleList.Count - 1)
                PlotConstRCircle(mlstRCircleList.Item(I), dblZo, Grobj)
            Next

            'Plot all extra Constant X Circle in the mlstRCircleList Array List.
            For I = 0 To (mlstXCircleList.Count - 1)
                PlotConstXCircle(mlstXCircleList.Item(I), dblZo, Grobj)
            Next

            'Plot all extra Constant G Circle in the mlstRCircleList Array List.
            For I = 0 To (mlstGCircleList.Count - 1)
                PlotConstGCircle(mlstGCircleList.Item(I), dblZo, Grobj)
            Next

            'Plot all extra Constant B Circle in the mlstRCircleList Array List.
            For I = 0 To (mlstBCircleList.Count - 1)
                PlotConstBCircle(mlstBCircleList.Item(I), dblZo, Grobj)
            Next

            'Draw at Constant R and X circles at the current cursor location if
            'Dynamic Z Coordinate option is enabled.
            If DynamicZCoorMenuItem.Checked = True Then
                PlotConstRCircle(mobjZ.Real, dblZo, Grobj)
                PlotConstXCircle(mobjZ.Imagninary, dblZo, Grobj)
            End If

            'Draw at Constant G and B circles at the current cursor location if
            'Dynamic Y Coordinate option is enabled.
            If DynamicYCoorMenuItem.Checked = True Then
                PlotConstGCircle(mobjY.Real, dblZo, Grobj)
                PlotConstBCircle(mobjY.Imagninary, dblZo, Grobj)
            End If

            'Draw a point at the cursor head, only do this if Dynamic Z or Y Coordinate option is enabled.
            If (DynamicZCoorMenuItem.Checked = True) Or (DynamicYCoorMenuItem.Checked = True) Then
                pt = ConvertReflectionCoeffToWinCoordinate(mobjT)
                pt.X = pt.X - (intPW / 2) 'To adjust the center position when point weigth is >1.
                pt.Y = pt.Y - (intPW / 2)
                Grobj.FillEllipse(mPointbrush, pt.X, pt.Y, mobjInterFormMessage.intPointWeigth, mobjInterFormMessage.intPointWeigth)
            End If

        Catch theException As Exception

            MessageBox.Show("MainPicBox_Paint " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub SmithChartForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        mobjInterFormMessage.intID = 0 'Indicate that Smith Chart resources are disposed.

    End Sub

    'Author:        Fabian Kung Wai Lee
    'Last modified: 16th Sep 2007
    'Description:   Subroutine to draw a circle corresponding to constant resistance (R)
    '               circle
    'Argument: 
    'dblR = Resistance
    'Zo = Reference impedance
    'objG = Graphic object
    Private Sub PlotConstRCircle(ByVal dblR As Double, ByVal Zo As Double, ByRef objG As Graphics)

        Dim CornerPoint As Point
        Dim intRad As Integer 'Radius of the circle
        Dim dblRp As Double
        Dim centerpoint As Point 'Center point of the circle

        Try
            'Find center and radius
            dblRp = dblR / Zo   'Normalize first.

            'Check for overflow
            If dblRp = -1 Then
                dblRp = -0.995
            End If

            intRad = Fix((1 / (1 + dblRp)) * mintSpan)
            If intRad < 0 Then
                intRad = -intRad 'only use positive value for radius
            End If
            centerpoint.X = ((dblRp / (1 + dblRp))) * mintSpan
            centerpoint.Y = 0

            'Draw circle
            CornerPoint = ConvertToWindowsCoordinate(centerpoint.X - intRad, centerpoint.Y + intRad)
            objG.DrawEllipse(mLinePen1, CornerPoint.X, CornerPoint.Y, 2 * intRad, 2 * intRad)

        Catch theException As Exception

            MessageBox.Show("PlotConstRCircle " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    'Author:        Fabian Kung Wai Lee
    'Last modified: 16th Sep 2007
    'Description:   Subroutine to draw a circle corresponding to constant reactance (X)
    '               circle
    'Argument: dblX = Reactance, 
    'Zo = Reference impedance
    'objG = Graphic object
    Private Sub PlotConstXCircle(ByVal dblX As Double, ByVal zo As Double, ByRef objG As Graphics)
        Dim CornerPoint As Point
        Dim sngRad As Single 'Radius of the circle
        Dim sngXp As Single
        Dim sngUo As Single
        Dim sngb As Single
        Dim sngalpha_deg As Single
        Dim sngStartAngle As Single
        Dim sngSweepAngle As Single
        Dim centerpoint As Point 'Center point of the circle

        Try
            'Check for overflow
            If dblX = 0 Then
                dblX = 0.001
            End If

            'Find center and radius
            sngXp = CSng(dblX / zo) 'Normalize first.
            sngRad = (1 / sngXp) * mintSpan
            If sngRad < 0 Then
                sngRad = -sngRad 'only use positive value for radius
            End If

            sngb = 1 / sngXp

            'X coordinate of the Intersecting point between constant X circle and
            'Smith Chart boundary (circle of unity radius)
            sngUo = (1 - (sngb * sngb)) / (1 + (sngb * sngb))

            'compute the angle alpha, this is dependent on whether X is positive or negative,
            'and the magnitude of X, i.e. |X|>Zo or |X|=<Zo.
            If sngXp > 0 Then 'positive X
                sngStartAngle = 90
                If sngXp > 1 Then
                    sngalpha_deg = (Atan(sngb * sngUo / (1 - sngUo)) * 180) / PI 'The angle in degrees
                    sngSweepAngle = sngalpha_deg + 90
                Else
                    sngalpha_deg = (Acos(-sngUo) * 180) / PI 'The angle in degrees
                    sngSweepAngle = sngalpha_deg
                End If
            Else 'negative X
                If sngXp < -1 Then
                    sngalpha_deg = (Acos((1 - sngUo) / (-sngb)) * 180) / PI 'The angle in degrees
                    sngStartAngle = 180 - sngalpha_deg
                    sngSweepAngle = 90 + sngalpha_deg
                Else
                    sngalpha_deg = (Acos((1 - sngUo) / (-sngb)) * 180) / PI 'The angle in degrees
                    sngStartAngle = 180 + sngalpha_deg
                    sngSweepAngle = 90 - sngalpha_deg
                End If
            End If

            centerpoint.X = (1) * mintSpan
            centerpoint.Y = (1 / sngXp) * mintSpan

            'Draw an arc corresponding to constant X circle
            CornerPoint = ConvertToWindowsCoordinate(CInt(centerpoint.X - sngRad), CInt(centerpoint.Y + sngRad))
            objG.DrawArc(mLinePen1, CSng(CornerPoint.X), CSng(CornerPoint.Y), (2 * sngRad), _
            (2 * sngRad), CSng(sngStartAngle), CSng(sngSweepAngle))

        Catch theException As Exception

            MessageBox.Show("PlotConstXCircle " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    'Author:        Fabian Kung Wai Lee
    'Last modified: 16th Sep 2007
    'Description:   Subroutine to draw a circle corresponding to constant conductance (G)
    '               circle
    'Argument: 
    'dblG = Conductance
    'Zo = Reference impedance
    'objG = Graphic object
    Private Sub PlotConstGCircle(ByVal dblG As Double, ByVal Zo As Double, ByRef objG As Graphics)

        Dim CornerPoint As Point
        Dim intRad As Integer 'Radius of the circle
        Dim dblGp As Double
        Dim centerpoint As Point 'Center point of the circle

        Try
            'Find center and radius
            dblGp = dblG * Zo   'Normalize first.

            'Check for overflow
            'If ((dblGp + 1) < 0.001) or ((dblGp+1)>0.001) Then
            'dblGp = -0.999
            'End If

            intRad = Fix((1 / (1 + dblGp)) * mintSpan)
            If intRad < 0 Then
                intRad = -intRad 'only use positive value for radius
            End If
            centerpoint.X = (-(dblGp / (1 + dblGp))) * mintSpan
            centerpoint.Y = 0

            'Draw circle
            CornerPoint = ConvertToWindowsCoordinate(centerpoint.X - intRad, centerpoint.Y + intRad)
            objG.DrawEllipse(mLinePen2, CornerPoint.X, CornerPoint.Y, 2 * intRad, 2 * intRad)

        Catch theException As Exception

                MessageBox.Show("PlotConstGCircle " & theException.Message, "Error", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    'Author:        Fabian Kung Wai Lee
    'Last modified: 16th Sep 2007
    'Description:   Subroutine to draw a circle corresponding to constant susceptance (B)
    '               circle
    'Argument: dblB = Susceptance, 
    'Zo = Reference impedance
    'objG = Graphic object
    Private Sub PlotConstBCircle(ByVal dblB As Double, ByVal zo As Double, ByRef objG As Graphics)
        Dim CornerPoint As Point
        Dim sngRad As Single 'Radius of the circle
        Dim sngBp As Single
        Dim sngVo As Single
        Dim sngc As Single
        Dim sngalpha_deg As Single
        Dim sngStartAngle As Single
        Dim sngSweepAngle As Single
        Dim centerpoint As Point 'Center point of the circle

        Try
            'Check for overflow
            If dblB = 0 Then
                dblB = 0.00001
            End If

            'Find center and radius
            sngBp = CSng(dblB * zo) 'Normalize first.
            sngRad = (1 / sngBp) * mintSpan
            If sngRad < 0 Then
                sngRad = -sngRad 'only use positive value for radius
            End If

            sngc = 1 / sngBp

            'Y coordinate of the Intersecting point between constant B circle and
            'Smith Chart boundary (circle of unity radius)
            sngVo = (-2 * sngc) / (1 + (sngc * sngc))

            'compute the angle alpha, this is dependent on whether B is positive or negative,
            'and the magnitude of B, i.e. |B|>Yo or |B|=<Yo.
            If sngBp > 0 Then 'Positive B
                If sngBp > 1 Then
                    sngalpha_deg = (Acos(-sngVo) * 180) / PI 'The angle in degrees
                    sngStartAngle = -90
                    sngSweepAngle = 90 + sngalpha_deg
                Else
                    sngalpha_deg = (Acos(-sngVo) * 180) / PI 'The angle in degrees
                    sngStartAngle = -90
                    sngSweepAngle = 90 - sngalpha_deg
                End If
            Else 'Negative B
                If sngBp < -1 Then
                    sngalpha_deg = (Acos(sngVo) * 180) / PI 'The angle in degrees
                    sngStartAngle = -sngalpha_deg
                    sngSweepAngle = 90 + sngalpha_deg
                Else
                    sngalpha_deg = (Acos(sngVo) * 180) / PI 'The angle in degrees
                    sngStartAngle = sngalpha_deg
                    sngSweepAngle = 90 - sngalpha_deg
                End If
            End If

            centerpoint.X = (-1) * mintSpan
            centerpoint.Y = (-1 / sngBp) * mintSpan

            'Draw an arc corresponding to constant B circle
            CornerPoint = ConvertToWindowsCoordinate(CInt(centerpoint.X - sngRad), CInt(centerpoint.Y + sngRad))
            objG.DrawArc(mLinePen2, CSng(CornerPoint.X), CSng(CornerPoint.Y), (2 * sngRad), _
            (2 * sngRad), CSng(sngStartAngle), CSng(sngSweepAngle))

        Catch theException As Exception

            MessageBox.Show("PlotConstBCircle " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    'Author:        Fabian Kung Wai Lee
    'Last modified: 4th Nov 2006
    'Description: Function to convert from a point from normal XY coordinate to relative 
    '             Windows coordinate system.  This coordinate has its XY origin
    '             at the upper left-hand corner of a graphic window.  A conventional XY
    '             XY coordinate system has the XY origin at the center of the graphic 
    '             window.
    'Argument: intX = Absolute X position
    '          intY = Absolute Y position
    'Return: an object of Point
    Private Function ConvertToWindowsCoordinate(ByVal intX As Integer, ByVal intY As Integer) _
    As Point
        Try
            ConvertToWindowsCoordinate.X = (mintOXw + mintOXwOffset) + intX
            ConvertToWindowsCoordinate.Y = (mintOYw + mintOYwOffset) - intY

        Catch theException As Exception
            MessageBox.Show("ConvertToWindowsCoordinate" & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    'Author:        Fabian Kung Wai Lee
    'Last modified: 4th Nov 2006
    'Description: Subroutine to convert from a point's Windows XY coordinate to reflection 
    '             coefficient.    
    'Argument: intX = X position in Windows coordinate system.
    '          intY = Y position in Windows coordinate system.
    'Return: None, the new reflection coefficient value is the global variable mobjT.
    Private Sub ConvertWinCoordinateToReflectionCoeff(ByVal intX As Integer, ByVal intY As Integer)
        Dim dblRe As Double
        Dim dblIm As Double

        Try
            dblRe = CDbl((intX - (mintOXw + mintOXwOffset)) / mintSpan)
            dblIm = CDbl((mintOYw + mintOYwOffset - intY) / mintSpan)
            mobjT.Real = dblRe
            mobjT.Imagninary = dblIm

        Catch theException As Exception
            MessageBox.Show("ConvertWinCoordinateToReflectionCoeff" & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    'Author:        Fabian Kung Wai Lee
    'Last modified: 3rd July 2007
    'Description: Subroutine to convert from a reflection coefficient to Windows coordinate
    '             system.
    'Argument: objT = a ComplexNumberClass object.
    'Return: Point object.
    Private Function ConvertReflectionCoeffToWinCoordinate(ByRef objT As ComplexNumberClass) As Point
        Dim intX As Integer
        Dim intY As Integer

        Try
            intX = (objT.Real) * mintSpan
            intY = objT.Imagninary * mintSpan
            ConvertReflectionCoeffToWinCoordinate.X = (mintOXw + mintOXwOffset) + intX
            ConvertReflectionCoeffToWinCoordinate.Y = (mintOYw + mintOYwOffset) - intY

        Catch theException As Exception
            MessageBox.Show("ConvertReflectionCoeffToWinCoordinate" & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    'Author:        Fabian Kung Wai Lee
    'Last modified: 4th Nov 2006
    'Description: This MouseWheel routine is used to Zoom into and out of the Smith Chart
    Private Sub SmithChartForm_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseWheel
        Dim intDelta As Integer
        Dim intCursorX As Integer
        Dim intCursorY As Integer

        Try
            'Get the (x,y) coordinate of the cursor relative to the main picturebox.
            intCursorX = e.X - mintPicBoxStartX
            intCursorY = e.Y - mintPicBoxStartY

            'Check if the cursor falls within the region of the main picturebox, if 
            'no ignore this event.
            If (intCursorX < 0) Or (intCursorX > MainPicBox.Width) Then
                Exit Sub
            End If
            If (intCursorY < 0) Or (intCursorY > MainPicBox.Height) Then
                Exit Sub
            End If

            'Change in span is determined by intDelta.  The larger the value of intDelta,
            'the smaller will be the change in span when user rotates the mouse's 
            'wheel.
            intDelta = 30
            mintDeltaSpan = mintSpan / intDelta

            'Set minimum value.
            If mintDeltaSpan = 0 Then
                mintDeltaSpan = 1
            End If
            If e.Delta > 0 Then
                mintSpan = mintSpan + mintDeltaSpan 'Increase the new Smith Chart radius
            Else
                mintSpan = mintSpan - mintDeltaSpan 'Decrease the new Smith Chart radius
            End If

            If mintSpan < 71 Then 'Fix the minimum size of the radius of the Smith Chart
                mintSpan = 71
            End If

            MainPicBox.Refresh() 'Redraw picture box content

        Catch theException As Exception

            MessageBox.Show("SmithChartForm_MouseWheel " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub SmithHScrollBar_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles SmithHScrollBar.Scroll
        Static intCount As Integer = 0

        Try
            'Check whether to scroll the Smith Chart to the right or left
            If (SmithHScrollBar.Value > mintHSBarOldValue) Then
                mintOXwOffset = mintOXwOffset + 2
                intCount = 0 'Reset counter
            ElseIf (SmithHScrollBar.Value < mintHSBarOldValue) Then
                mintOXwOffset = mintOXwOffset - 2
                intCount = 0 'Reset counter
            ElseIf (SmithHScrollBar.Value = 0) Then
                'This indicate the min. end of the scrollbar is reached
                mintOXwOffset = mintOXwOffset - 4
                intCount = 0 'Reset counter
            ElseIf (SmithHScrollBar.Value = mintHSBarOldValue) Then
                'There are two instances when this condition occur. When the user 
                'clicks the scrollbar's increment or decrement button (with a mouse)
                'and when the maximum position is reached.  When the user clicks the 
                'increment or decrement button, the Value property of the control still 
                'retain the old value. Thus this event trigger.  It is after the user
                'releases the mouse button that the Value property changes. I discovered 
                'that the Value property of the scrollbar control cannot reach the 
                'maximum value as specified in the Maximum property. Thus to detect 
                'the maximum position reached condition, intCount is used.  The variable
                'intCount will be incremented at least twice under maximum position
                'reached condition.
                intCount = intCount + 1
                If intCount > 1 Then
                    mintOXwOffset = mintOXwOffset + 4
                    intCount = 0 'Reset counter
                End If
            End If
            mintHSBarOldValue = SmithHScrollBar.Value 'Update the current value

            MainPicBox.Refresh() 'Redraw picture box content

        Catch theException As Exception

            MessageBox.Show("SmithHScrollBar_Scroll " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub SmithVScrollBar_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles SmithVScrollBar.Scroll
        Static intCount As Integer = 0

        Try
            'Check whether to scroll the Smith Chart up or down
            If (SmithVScrollBar.Value > mintVSBarOldValue) Then
                mintOYwOffset = mintOYwOffset + 2
                intCount = 0 'Reset counter
            ElseIf (SmithVScrollBar.Value < mintVSBarOldValue) Then
                mintOYwOffset = mintOYwOffset - 2
                intCount = 0 'Reset counter
            ElseIf (SmithVScrollBar.Value = 0) Then
                'This indicate the min. end of the scrollbar is reached
                mintOYwOffset = mintOYwOffset - 4
                intCount = 0 'Reset counter
            ElseIf (SmithVScrollBar.Value = mintVSBarOldValue) Then
                'There are two instances when this condition occur. When the user 
                'clicks the scrollbar's increment or decrement button (with a mouse)
                'and when the maximum position is reached.  When the user clicks the 
                'increment or decrement button, the Value property of the control still 
                'retain the old value. Thus this event trigger.  It is after the user
                'releases the mouse button that the Value property changes. I discovered 
                'that the Value property of the scrollbar control cannot reach the 
                'maximum value as specified in the Maximum property. Thus to detect 
                'the maximum position reached condition, intCount is used.  The variable
                'intCount will be incremented at least twice under maximum position
                'reached condition.
                intCount = intCount + 1
                If intCount > 1 Then
                    mintOYwOffset = mintOYwOffset + 4
                    intCount = 0 'Reset counter
                End If
            End If
            mintVSBarOldValue = SmithVScrollBar.Value 'Update the current value

            MainPicBox.Refresh() 'Redraw picture box content
        Catch theException As Exception

            MessageBox.Show("SmithVScrollBar_Scroll " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub FitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FitButton.Click

        Try
            'Determine the size of the Smith Chart and its origin with respect
            'to the picture box
            'Finding the radius:
            If MainPicBox.Width > MainPicBox.Height Then
                mintSpan = (MainPicBox.Height / 2) - 4 'allow some clearance from the side
            Else
                mintSpan = (MainPicBox.Width / 2) - 4 'allow some clearance from the side
            End If

            'Setting the origin of the main picturebox coordinate system
            mintOXw = MainPicBox.Width / 2
            mintOYw = MainPicBox.Height / 2
            mintOXwOffset = 0
            mintOYwOffset = 0

            MainPicBox.Refresh() 'Redraw picture box content

        Catch theException As Exception

            MessageBox.Show("FitButton_Click" & theException.Message, "Error", _
                      MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub ZCoorCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZCoorCheckBox.CheckedChanged
        MainPicBox.Refresh()    'Redraw MainPicBox.
    End Sub

    Private Sub YCoorCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YCoorCheckBox.CheckedChanged
        MainPicBox.Refresh()    'Redraw MainPicBox.
    End Sub

    Private Sub InputModeComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputModeComboBox.SelectedIndexChanged
        Dim intItem As Integer

        intItem = InputModeComboBox.SelectedIndex

        If intItem = 0 Then
            mintInputMode = 0 'Z input mode.
            Input1Label.Text = "R"
            Input2Label.Text = "X"
        ElseIf intItem = 1 Then
            mintInputMode = 1 'Y input mode.
            Input1Label.Text = "G"
            Input2Label.Text = "B"
        Else
            mintInputMode = 2 'Rho input mode.
            Input1Label.Text = "Re(T)"
            Input2Label.Text = "Im(T)"
        End If
    End Sub

    Private Sub PlotButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlotButton.Click

        Dim dblRe As Double
        Dim dblIm As Double
        Dim objZ As New ComplexNumberClass(0, 0)
        Dim objY As New ComplexNumberClass(0, 0)

        Try

            If Input1TextBox.Text <> "" Then   'Make sure it is not empty
                dblRe = CDbl(Input1TextBox.Text)  'Make sure the value is between 0 and 255.
            Else
                dblRe = 0
                Input1TextBox.Text = "0"
            End If
            If Input2TextBox.Text <> "" Then   'Make sure it is not empty
                dblIm = CDbl(Input2TextBox.Text)  'Make sure the value is between 0 and 255.
            Else
                dblIm = 0
                Input2TextBox.Text = "0"
            End If

            'Update master index first.
            If mintCount < MAXCOUNT Then
                mintCount = mintCount + 1 'Increment reflection coefficient array pointer.
            Else
                mintCount = 0   'Reset reflection coefficent array pointer.
            End If

            'Capture value keyed in by user.
            If mintInputMode = 0 Then 'Z input mode
                objZ.Real = dblRe    'Create a complex impedance Z from user inputs.
                objZ.Imagninary = dblIm
                mobjTo(mintCount) = ZtoT(objZ) 'Convert Z into reflection coefficient T.
            ElseIf mintInputMode = 1 Then 'Y input mode
                objY.Real = dblRe    'Create a complex admittance Y from user inputs.
                objY.Imagninary = dblIm
                objZ = objY.DivComplex(mobj1, objY) 'Convert Y into Z.
                mobjTo(mintCount) = ZtoT(objZ) 'Convert Z into reflection coefficient T.
            Else 'Rho input mode
                mobjTo(mintCount).Real = dblRe
                mobjTo(mintCount).Imagninary = dblIm
            End If

            mobjTo(mintCount).Visible = True   'Enable this reflection coefficient to be shown.
            MainPicBox.Refresh()    'Redraw MainPicBox.

        Catch theException As Exception
            MessageBox.Show("PlotButton_Click" & theException.Message, "Error", _
                      MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub ClearButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearButton.Click

        Dim I As Integer

        For I = 0 To MAXCOUNT
            mobjTo(I).Visible = False   'Set visibility property to false for all reflection coefficients.
        Next
        MainPicBox.Refresh()    'Redraw MainPicBox.

    End Sub

    'Author:        Fabian Kung Wai Lee
    'Last modified: 1st July 2007
    'Description:   This function is to convert reflection coefficient (T) into
    '               impedance (Z).
    'Argument: objT = an object of ComplexNumberClass
    'Return: The impedance value in ComplexNumberClass
    Private Function TtoZ(ByRef objT As ComplexNumberClass) As ComplexNumberClass

        Dim objZ As New ComplexNumberClass(0, 0)
        'Declare a complex value: objZo = Zo + j0, Zo = reference impedance.
        Dim objZo As New ComplexNumberClass(mobjInterFormMessage.dblZo, 0)

        Try

            'We use this formula: Z = ((1+T)/(1-T))Zo, 
            'where T = reflection coefficient and
            'Zo = reference impedance
            objZ = objZ.AddComplex(mobj1, objT) '1+T
            objZ = objZ.DivComplex(objZ, objZ.SubComplex(mobj1, objT)) '(1+T)/(1-T)
            objZ = objZ.MulComplex(objZ, objZo) 'Zo*((1+T)/(1-T))
            TtoZ = objZ

        Catch theException As Exception
            MessageBox.Show("TtoZ" & theException.Message, "Error", _
                      MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    'Author:        Fabian Kung Wai Lee
    'Last modified: 3rd July 2007
    'Description:   This function is to convert impedance (Z) into reflection 
    '               coefficient (T).
    'Argument: objZ = an object of ComplexNumberClass
    'Return: The reflection coefficient value in ComplexNumberClass
    Private Function ZtoT(ByRef objZ As ComplexNumberClass) As ComplexNumberClass

        Dim objC As New ComplexNumberClass(0, 0)
        'Declare a complex value: objZo = Zo + j0, Zo = reference impedance.
        Dim objZo As New ComplexNumberClass(mobjInterFormMessage.dblZo, 0)

        objC = objZ.SubComplex(objZ, objZo)
        ZtoT = objC.DivComplex(objC, objC.AddComplex(objZ, objZo))
    End Function

    Private Sub RCircleMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RCircleMenuItem.Click

        'Adds a new item into the resistance value list, up to 6 items.
        If mlstRCircleList.Count < 7 Then
            mlstRCircleList.Add(CDbl(RLabel.Text))
            MainPicBox.Refresh()    'Redraw MainPicBox.
        End If

    End Sub

    Private Sub XCircleMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XCircleMenuItem.Click

        'Adds a new item into the reactance value list, up to 6 items.
        If mlstXCircleList.Count < 7 Then
            mlstXCircleList.Add(CDbl(XLabel.Text))
            MainPicBox.Refresh()    'Redraw MainPicBox.
        End If

    End Sub

    Private Sub GCircleMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GCircleMenuItem.Click

        'Adds a new item into the conductance value list, up to 6 items.
        If mlstGCircleList.Count < 7 Then
            mlstGCircleList.Add(CDbl(GLabel.Text))
            MainPicBox.Refresh()    'Redraw MainPicBox.
        End If

    End Sub

    Private Sub BCircleMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BCircleMenuItem.Click

        'Adds a new item into the susceptance value list, up to 6 items.
        If mlstBCircleList.Count < 7 Then
            mlstBCircleList.Add(CDbl(BLabel.Text))
            MainPicBox.Refresh()    'Redraw MainPicBox.
        End If

    End Sub

    Private Sub PlotMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlotMenuItem.Click

        Try
            'Update master index first.
            If mintCount < MAXCOUNT Then
                mintCount = mintCount + 1 'Increment reflection coefficient array pointer.
            Else
                mintCount = 0   'Reset reflection coefficent array pointer.
            End If

            'Capture value at cursor's position.
            mobjTo(mintCount).Real = mobjT.Real
            mobjTo(mintCount).Imagninary = mobjT.Imagninary

            mobjTo(mintCount).Visible = True   'Enable this reflection coefficient to be shown.
            MainPicBox.Refresh()    'Redraw MainPicBox.

        Catch theException As Exception
            MessageBox.Show("PlotMenuItem_Click" & theException.Message, "Error", _
                      MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub ClearMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearMenuItem.Click

        mlstRCircleList.Clear() 'Clears all items in the Constant R Circles Array List.
        mlstXCircleList.Clear() 'Clears all items in the Constant X Circles Array List.
        mlstGCircleList.Clear() 'Clears all items in the Constant G Circles Array List.
        mlstBCircleList.Clear() 'Clears all items in the Constant B Circles Array List.
        MainPicBox.Refresh()    'Redraw MainPicBox.

    End Sub

    Private Sub DynamicZCoorMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DynamicZCoorMenuItem.Click

        If DynamicZCoorMenuItem.Checked = False Then
            DynamicZCoorMenuItem.Checked = True
        Else
            DynamicZCoorMenuItem.Checked = False
        End If
        MainPicBox.Refresh()    'Redraw MainPicBox.

    End Sub

    Private Sub DynamicYCoorMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DynamicYCoorMenuItem.Click

        If DynamicYCoorMenuItem.Checked = False Then
            DynamicYCoorMenuItem.Checked = True
        Else
            DynamicYCoorMenuItem.Checked = False
        End If
        MainPicBox.Refresh()    'Redraw MainPicBox.

    End Sub
End Class
