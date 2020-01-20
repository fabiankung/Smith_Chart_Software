'Author:            Fabian Kung Wai Lee
'Last modified:     7th July 2007
'Description:       Inter modules communication message class   
'Filename:          MessageClass.vb
'Language:          Visual Basic .NET
'Development tool:  Microsoft Visual Studio .NET 2003

Public Class MessageClass

    Public Shared dblZo As Double
    Public Shared dblFreq As Double
    Public Shared intID As Integer              'ID to indicate existent of SmithChart window, 1 = class exist, 0 = class has been cleared from memory.
    Public Shared intID2 As Integer             'ID to indicate existent of ImpedanceTransformForm window
    Public Shared intID3 As Integer             'ID to indicate existent of ImpedanceTransformForm2 window
    Public Shared backgroundcolor As Color
    Public Shared intLineThickness As Integer
    Public Shared Linecolor1 As Color
    Public Shared Linecolor2 As Color
    Public Shared Pointcolor As Color
    Public Shared intPointWeigth As Integer

    Public Sub New()

        MyBase.new()

    End Sub

End Class
