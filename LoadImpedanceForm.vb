Public Class ZLForm
    '--- Module level declarations ---
    'Environment:
    Public mstrFormTitle As String
    Public mstrVersion As String
    Public mobjInterFormMessage As New MessageClass

    'Variables:
    Public mobjZL As New ComplexNumberClass(0, 0) 'Load impedance


    Private Sub ZLForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'Update the title and version of the main window
            Me.Text = mstrFormTitle & " " & "Version " & mstrVersion

        Catch theException As Exception

            MessageBox.Show("ImpedanceTransformForm_Load " & theException.Message, "Error", _
                     MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs) Handles OkButton.Click

        Dim dblRe As Double
        Dim dblIm As Double

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


            'Capture value keyed in by user.

            mobjZL.Real = dblRe    'Create a complex impedance Z from user inputs.
            mobjZL.Imagninary = dblIm
            Close()

        Catch theException As Exception
            MessageBox.Show("OkButton_Click" & theException.Message, "Error", _
                      MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Close()
    End Sub

End Class