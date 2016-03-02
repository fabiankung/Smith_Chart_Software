Public Class DisplaySettingForm
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
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BackGroundColorButton As System.Windows.Forms.Button
    Friend WithEvents DisplayColorSelectDialog As System.Windows.Forms.ColorDialog
    Friend WithEvents ChartPanel As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LineThicknessTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents PointWeightTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents PointColorButton As System.Windows.Forms.Button
    Friend WithEvents SettingCancelButton As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LineColor1Button As System.Windows.Forms.Button
    Friend WithEvents LineColor2Button As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DisplaySettingForm))
        Me.OKButton = New System.Windows.Forms.Button
        Me.BackGroundColorButton = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.LineThicknessTextBox = New System.Windows.Forms.TextBox
        Me.ChartPanel = New System.Windows.Forms.Panel
        Me.DisplayColorSelectDialog = New System.Windows.Forms.ColorDialog
        Me.SettingCancelButton = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.LineColor1Button = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.PointWeightTextBox = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.PointColorButton = New System.Windows.Forms.Button
        Me.LineColor2Button = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(120, 296)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(56, 40)
        Me.OKButton.TabIndex = 0
        Me.OKButton.Text = "&OK"
        '
        'BackGroundColorButton
        '
        Me.BackGroundColorButton.Location = New System.Drawing.Point(120, 16)
        Me.BackGroundColorButton.Name = "BackGroundColorButton"
        Me.BackGroundColorButton.Size = New System.Drawing.Size(32, 24)
        Me.BackGroundColorButton.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 144)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 32)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Set Line Thickness"
        '
        'LineThicknessTextBox
        '
        Me.LineThicknessTextBox.Location = New System.Drawing.Point(120, 152)
        Me.LineThicknessTextBox.Name = "LineThicknessTextBox"
        Me.LineThicknessTextBox.Size = New System.Drawing.Size(40, 20)
        Me.LineThicknessTextBox.TabIndex = 4
        Me.LineThicknessTextBox.Text = ""
        '
        'ChartPanel
        '
        Me.ChartPanel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.ChartPanel.Location = New System.Drawing.Point(64, 224)
        Me.ChartPanel.Name = "ChartPanel"
        Me.ChartPanel.Size = New System.Drawing.Size(80, 56)
        Me.ChartPanel.TabIndex = 5
        '
        'SettingCancelButton
        '
        Me.SettingCancelButton.Location = New System.Drawing.Point(48, 296)
        Me.SettingCancelButton.Name = "SettingCancelButton"
        Me.SettingCancelButton.Size = New System.Drawing.Size(56, 40)
        Me.SettingCancelButton.TabIndex = 6
        Me.SettingCancelButton.Text = "&Cancel"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 32)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Set Background Color"
        '
        'LineColor1Button
        '
        Me.LineColor1Button.Location = New System.Drawing.Point(120, 48)
        Me.LineColor1Button.Name = "LineColor1Button"
        Me.LineColor1Button.Size = New System.Drawing.Size(32, 24)
        Me.LineColor1Button.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Set Z line Color"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 192)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 24)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Point Weight"
        '
        'PointWeightTextBox
        '
        Me.PointWeightTextBox.Location = New System.Drawing.Point(120, 192)
        Me.PointWeightTextBox.Name = "PointWeightTextBox"
        Me.PointWeightTextBox.Size = New System.Drawing.Size(40, 20)
        Me.PointWeightTextBox.TabIndex = 11
        Me.PointWeightTextBox.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 24)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Set Point Color"
        '
        'PointColorButton
        '
        Me.PointColorButton.Location = New System.Drawing.Point(120, 112)
        Me.PointColorButton.Name = "PointColorButton"
        Me.PointColorButton.Size = New System.Drawing.Size(32, 24)
        Me.PointColorButton.TabIndex = 13
        '
        'LineColor2Button
        '
        Me.LineColor2Button.Location = New System.Drawing.Point(120, 80)
        Me.LineColor2Button.Name = "LineColor2Button"
        Me.LineColor2Button.Size = New System.Drawing.Size(32, 24)
        Me.LineColor2Button.TabIndex = 14
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 80)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Set Y Line Color"
        '
        'DisplaySettingForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(224, 350)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.LineColor2Button)
        Me.Controls.Add(Me.PointColorButton)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.PointWeightTextBox)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.LineColor1Button)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.SettingCancelButton)
        Me.Controls.Add(Me.ChartPanel)
        Me.Controls.Add(Me.LineThicknessTextBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BackGroundColorButton)
        Me.Controls.Add(Me.OKButton)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "DisplaySettingForm"
        Me.Text = "DisplaySettingForm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Module level declarations
    Public mstrFormTitle As String
    Public mstrVersion As String

    Public InterFormMessage As New MessageClass

    Private Linecolor1 As Color
    Private Linecolor2 As Color
    Private intLineThickness As Integer
    Private Pointcolor As Color
    Private intPointWeigth As Integer

    Private Sub DisplaySettingForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            'Update the title and version of the main window
            Me.Text = mstrFormTitle & " " & "Version " & mstrVersion

            Me.CenterToScreen() 'Put this form in the center of the screen

            'Initialize all environment parameters on the sample panel
            ChartPanel.BackColor = InterFormMessage.backgroundcolor
            Linecolor1 = InterFormMessage.Linecolor1
            Linecolor2 = InterFormMessage.Linecolor2
            Pointcolor = InterFormMessage.Pointcolor
            intLineThickness = InterFormMessage.intLineThickness
            LineThicknessTextBox.Text = CStr(InterFormMessage.intLineThickness)
            intPointWeigth = InterFormMessage.intPointWeigth
            PointWeightTextBox.Text = CStr(InterFormMessage.intPointWeigth)

        Catch

            MessageBox.Show("Error in initialization", "Error", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub BackGroundColorButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackGroundColorButton.Click

        DisplayColorSelectDialog.Color = ChartPanel.BackColor 'Reload original color
        DisplayColorSelectDialog.ShowDialog() 'Show modal dialog box
        ChartPanel.BackColor = DisplayColorSelectDialog.Color 'Assign new color

    End Sub

    Private Sub LineColor1Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LineColor1Button.Click

        DisplayColorSelectDialog.Color = Linecolor1 'Reload original color
        DisplayColorSelectDialog.ShowDialog() 'Show modal dialog box   
        Linecolor1 = DisplayColorSelectDialog.Color 'Assign new color
        ChartPanel.Refresh() 'Redraw the chart panel using new settings

    End Sub

    Private Sub LineColor2Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LineColor2Button.Click

        DisplayColorSelectDialog.Color = Linecolor2 'Reload original color
        DisplayColorSelectDialog.ShowDialog() 'Show modal dialog box   
        Linecolor2 = DisplayColorSelectDialog.Color 'Assign new color
        ChartPanel.Refresh() 'Redraw the chart panel using new settings

    End Sub

    Private Sub PointColorButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PointColorButton.Click

        DisplayColorSelectDialog.Color = Pointcolor 'Reload original color
        DisplayColorSelectDialog.ShowDialog() 'Show modal dialog box   
        Pointcolor = DisplayColorSelectDialog.Color 'Assign new color
        ChartPanel.Refresh() 'Redraw the chart panel using new settings

    End Sub

    Private Sub LineThicknessTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LineThicknessTextBox.TextChanged

        Try
            If LineThicknessTextBox.Text <> "" Then
                intLineThickness = CInt(LineThicknessTextBox.Text)
            End If
            ChartPanel.Refresh() 'Redraw the chart panel using new settings
        Catch
            MessageBox.Show("Please input a valid numeric value", "Error", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub PointWeightTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PointWeightTextBox.TextChanged

        Try
            If PointWeightTextBox.Text <> "" Then
                intPointWeigth = CInt(PointWeightTextBox.Text)
            End If
            ChartPanel.Refresh() 'Redraw the chart panel using new settings
        Catch
            MessageBox.Show("Please input a valid numeric value", "Error", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub ChartPanel_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles ChartPanel.Paint
        Dim ChartGraphics As Graphics = e.Graphics
        Dim LinePen1 As New Pen(Linecolor1, intLineThickness)
        Dim LinePen2 As New Pen(Linecolor2, intLineThickness)
        Dim PointBrush As New SolidBrush(Pointcolor)
        Dim intH As Integer = ChartPanel.Height
        Dim intW As Integer = ChartPanel.Width
        Dim StartPoint1 As New Point(intW / 2, 0)
        Dim EndPoint1 As New Point(intW / 2, intH)
        Dim StartPoint2 As New Point(0, intH / 2)
        Dim EndPoint2 As New Point(intW, intH / 2)

        'Draw a vertical and horizontal line in the ChartPanel
        ChartGraphics.DrawLine(LinePen1, StartPoint1, EndPoint1)
        ChartGraphics.DrawLine(LinePen2, StartPoint2, EndPoint2)
        'Draw a filled circle in the ChartPanel
        ChartGraphics.FillEllipse(PointBrush, 20, 20, intPointWeigth, intPointWeigth)

    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click

        'Assign new parameters to InterFormMessage class environment fields
        With InterFormMessage
            .backgroundcolor = ChartPanel.BackColor
            .Linecolor1 = Linecolor1
            .Linecolor2 = Linecolor2
            .intLineThickness = intLineThickness
            .Pointcolor = Pointcolor
            .intPointWeigth = intPointWeigth
        End With

        Close()

    End Sub

    Private Sub SettingCancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettingCancelButton.Click

        Close()

    End Sub

End Class
