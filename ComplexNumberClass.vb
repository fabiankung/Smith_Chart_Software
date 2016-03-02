'Author:            Fabian Kung Wai Lee
'Last modified:     3rd July 2007
'Description:       Complex number class, a type of the form a+jb, where
'                   j = square root of -1.
'Filename:          ComplexNumberClass.vb
'Language:          Visual Basic .NET
'Development tool:  Microsoft Visual Studio .NET 2003

Imports System.Math

Public Class ComplexNumberClass

    Private dblRe As Double 'Real part of the complex number
    Private dblIm As Double 'Imagninary part of the complex number
    'If blnVisible is assert true this complex number can be displayed,
    'say on a 2D XY plot or other format, such as Smith Chart.
    Private blnVisible As Boolean

    'Constructor
    Public Sub New(ByVal real As Double, ByVal img As Double)

        Me.dblRe = real
        Me.dblIm = img
        Me.blnVisible = False

    End Sub

    'Property  to access the complex number object values, i.e. real, imaginary,
    'magnitude and phase.
    'Real part:
    Public Property Real() As Double
        Get
            Return dblRe
        End Get
        Set(ByVal Value As Double)
            dblRe = Value
        End Set
    End Property

    'Imagnary part:
    Public Property Imagninary() As Double
        Get
            Return dblIm
        End Get
        Set(ByVal Value As Double)
            dblIm = Value
        End Set
    End Property

    'magnitude (readonly)
    Public ReadOnly Property Magnitude() As Double
        Get
            Return Sqrt((dblRe * dblRe) + (dblIm * dblIm))
        End Get
    End Property

    'Phase (readonly), in radian from 0 to 2*PI
    Public ReadOnly Property Phase() As Double
        Get
            Dim dblPhase As Double

            dblPhase = Atan(dblIm / dblRe)
            If dblIm > 0 Then
                If dblRe > 0 Then
                    Return dblPhase
                Else
                    Return (dblPhase + PI)
                End If
            Else
                If dblRe > 0 Then
                    Return (dblPhase + (2 * PI))
                Else
                    Return (dblPhase + PI)
                End If
            End If
        End Get
    End Property

    'Visible flag:
    Public Property Visible() As Boolean
        Get
            Return blnVisible
        End Get
        Set(ByVal Value As Boolean)
            blnVisible = Value
        End Set
    End Property

    'Complex addition method
    Function AddComplex(ByRef A As ComplexNumberClass, ByRef B As ComplexNumberClass) As ComplexNumberClass

        Dim objC As New ComplexNumberClass(0, 0)

        objC.dblRe = A.dblRe + B.dblRe
        objC.dblIm = A.dblIm + B.dblIm
        AddComplex = objC

    End Function

    'Complex subtraction method
    Function SubComplex(ByRef A As ComplexNumberClass, ByRef B As ComplexNumberClass) As ComplexNumberClass

        Dim objC As New ComplexNumberClass(0, 0)

        objC.dblRe = A.dblRe - B.dblRe
        objC.dblIm = A.dblIm - B.dblIm
        SubComplex = objC

    End Function

    'Complex conjugation method
    Function ConjComplex(ByVal A As ComplexNumberClass) As ComplexNumberClass

        ConjComplex.dblRe = A.dblRe
        ConjComplex.dblIm = -A.dblIm

    End Function

    'Complex multiplication method
    Function MulComplex(ByRef A As ComplexNumberClass, ByRef B As ComplexNumberClass) As ComplexNumberClass

        Dim objC As New ComplexNumberClass(0, 0)

        objC.dblRe = (A.dblRe * B.dblRe) - (A.dblIm * B.dblIm)
        objC.dblIm = (A.dblRe * B.dblIm) + (A.dblIm * B.dblRe)
        MulComplex = objC

    End Function

    'Complex division method
    Function DivComplex(ByRef A As ComplexNumberClass, ByRef B As ComplexNumberClass) As ComplexNumberClass

        Dim objC As New ComplexNumberClass(0, 0)
        Dim magBs As Double

        magBs = (B.dblRe) * (B.dblRe) + (B.dblIm) * (B.dblIm)
        magBs = 1 / magBs
        objC.dblRe = (magBs) * ((A.dblRe * B.dblRe) + (A.dblIm * B.dblIm))
        objC.dblIm = (magBs) * ((A.dblIm * B.dblRe) - (A.dblRe * B.dblIm))
        DivComplex = objC

    End Function
End Class
