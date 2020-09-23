VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Draw Cubic Spline"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Control Points"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   0
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Here is an absolute minimum Cubic Spline routine.
    
'It's a VB rewrite from a Java applet I found by by Anthony Alto 4/25/99

'Computes coefficients based on equations mathematically derived from the curve
'constraints.   i.e. :
'    curves meet at knots (predefined points)  - These must be sorted by X
'    first derivatives must be equal at knots
'    second derivatives must be equal at knots

Private Const nPoints = 6
Private x(nPoints) As Single
Private y(nPoints) As Single
Private p(nPoints) As Single    '
Private u(nPoints) As Single

Private Sub Command1_Click()
    Dim i As Integer
    x(1) = 10:      y(1) = 10
    x(2) = 50:      y(2) = 100
    x(3) = 100:     y(3) = 30
    x(4) = 150:     y(4) = 70
    x(5) = 200:     y(5) = 170
    x(6) = 250:     y(6) = 10
    For i = 1 To nPoints
        Picture1.Circle (x(i), y(i)), 4, &HFFFFFF
    Next
    Command2.Enabled = True
End Sub


Private Sub Command2_Click()
    Dim piece As Integer, xPos As Single, yPos As Single
   
    Call SetPandU
    
    For piece = 1 To nPoints - 1
        For xPos = x(piece) To x(piece + 1)
            yPos = getCurvePoint(piece, xPos)
            Picture1.PSet (xPos, yPos), &H0
        Next
    Next
End Sub

Private Function getCurvePoint(i As Integer, v As Single) As Single
    Dim t As Single
    'derived curve equation (which uses p's and u's for coefficients)
    t = (v - x(i)) / u(i)
    getCurvePoint = t * y(i + 1) + (1 - t) * y(i) + u(i) * u(i) * (F(t) * p(i + 1) + F(1 - t) * p(i)) / 6#
End Function

Private Function F(x As Single) As Single
        F = x * x * x - x
End Function

Private Sub SetPandU()
    Dim i As Integer
    Dim d(nPoints) As Single
    Dim w(nPoints) As Single
'Routine to compute the parameters of our cubic spline.  Based on equations derived from some basic facts...
'Each segment must be a cubic polynomial.  Curve segments must have equal first and second derivatives
'at knots they share.  General algorithm taken from a book which has long since been lost.

'The math that derived this stuff is pretty messy...  expressions are isolated and put into
'arrays.  we're essentially trying to find the values of the second derivative of each polynomial
'at each knot within the curve.  That's why theres only N-2 p's (where N is # points).
'later, we use the p's and u's to calculate curve points...

    For i = 2 To nPoints - 1
        d(i) = 2 * (x(i + 1) - x(i - 1))
    Next
    For i = 1 To nPoints - 1
        u(i) = x(i + 1) - x(i)
    Next
    For i = 2 To nPoints - 1
        w(i) = 6# * ((y(i + 1) - y(i)) / u(i) - (y(i) - y(i - 1)) / u(i - 1))
    Next
    For i = 2 To nPoints - 2
        w(i + 1) = w(i + 1) - w(i) * u(i) / d(i)
        d(i + 1) = d(i + 1) - u(i) * u(i) / d(i)
    Next
    p(1) = 0#
    For i = nPoints - 1 To 2 Step -1
        p(i) = (w(i) - u(i) * p(i + 1)) / d(i)
    Next
    p(nPoints) = 0#
End Sub
