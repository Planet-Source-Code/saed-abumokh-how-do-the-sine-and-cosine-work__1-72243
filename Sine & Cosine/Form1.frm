VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Ellipse"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y
Private Function Cosine(Number As Double) As Double

'                x^2    x^4    x^6    x^8    x^10   x^12   x^14  x^16
'   cos(x)= 1 - ---- + ---- - ---- + ---- - ---- + ---- - ---- + ---- ......
'                2!     4!     6!     8!     10!    12!    14!   16!
    
    Static PlusMinus As Boolean
    Dim i As Double
    '                       A                A is the upper
    Dim Upper As Double ' -----
    '                       B                B is the lower
    Dim Lower As Double
    
    Cosine = 1
    
    PlusMinus = False
     ' the cosine is infinety function, so i used the For from 2 to a big number
     ' you can change this number but not so big that make the function slow, and not lower than 20
     ' the number should be even number (2,4,6,8,10,12,14,16,18,20....)
    For i = 2 To 150 Step 2
        Upper = Number ^ i
        Lower = Factorial(i)
        PlusMinus = Not (PlusMinus)
        If PlusMinus = True Then
            Cosine = Cosine - (Upper / Lower)
        ElseIf PlusMinus = False Then
            Cosine = Cosine + (Upper / Lower)
        End If
    Next
End Function

Private Function Cosine2(Number As Double) As Double

    '                  Pi
    'Cos(x) = Sin(x + ---- )
    '                  2
    
    Dim Pi: Pi = "3.14159265358979323846264338327950288419716939937510582097494459230781640628620899862803482534211706798214808651328230664709384460955058223172535940812848111745028410270192852110555964462294895493038196442881097566593344612847564823378678316527120190914564856692346034861045432664821339360726024914127372458700660631558817488152092096282925409171536436789259036001133055305488204665213841469519415116094330572703657595919530921861173817326117931051185480744623799627495673518857527248912279381830119491298336733624406566430"
    Cosine2 = Sine(Number + (Pi / 2))
End Function
Private Function Sine(Number As Double) As Double
    
'                x^3    x^5    x^7    x^9    x^11   x^13   x^15  x^17
'   sin(x)= x - ---- + ---- - ---- + ---- - ---- + ---- - ---- + ---- .....
'                3!     5!     7!     9!     11!    13!    15!   17!
    
    Static PlusMinus As Boolean
    Dim i As Double
    Dim Upper As Double
    Dim Lower As Double
    
    Sine = Number
    
    PlusMinus = False
     '  the sine is infinety function, so i used the For from 3 to a big number
     ' you can change this number but not so big that make the function slow, and not lower than 20
     ' the number should be odd number (3,5,7,9,11,13,15,17,19,21....)
    For i = 3 To 151 Step 2
        Upper = Number ^ i
        Lower = Factorial(i)
        PlusMinus = Not (PlusMinus)
        If PlusMinus = True Then
            Sine = Sine - (Upper / Lower)
        ElseIf PlusMinus = False Then
            Sine = Sine + (Upper / Lower)
        End If
    Next
End Function

Private Function Factorial(Number As Double) As Double
    'Factorial(x) = x! = 1 * 2 * 3 * 4 * 5 * 6 ...... * x
    Factorial = 1
    For i = 1 To Number
        Factorial = Factorial * i
    Next
End Function

Private Sub Command1_Click()
    Me.CurrentX = 300
    Me.CurrentY = 400

    Timer1 = True
End Sub

Private Sub Timer1_Timer()
    x = x + 5
    x = x Mod 360
    Me.Line -(Sine(x * 3.141 / 180) * 200 + 300, Cosine(x * 3.141 / 180) * 100 + 300)
    Me.Line -(Sin(x * 3.141 / 180) * 200 + 300, Cos(x * 3.141 / 180) * 100 + 300)
End Sub
