VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "BallSim Beta - Scott Buckley 2004 (Click to add)"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Balls As Integer
Const Rad As Single = 300
Dim Ang As Double
Dim Xp(1000) As Single
Dim Yp(1000) As Single
Dim Cx As Double
Dim Cy As Double
Dim Clr(1000) As Long

Private Sub Form_Load()
    Xp(0) = ScaleWidth / 2
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Balls = Balls + 1
    Xp(Balls) = x
    Yp(Balls) = Y
    Clr(Balls) = Rnd * vbWhite
End Sub

Private Sub Timer1_Timer()
    For x = 0 To Balls
        Yp(x) = Yp(x) + 45
        If Yp(x) > ScaleHeight - Rad Then Yp(x) = ScaleHeight - Rad
        If Xp(x) < Rad Then Xp(x) = Rad
        If Xp(x) > ScaleWidth - Rad Then Xp(x) = ScaleWidth - Rad
        For Y = 0 To Balls
            If Y <> x Then
                If ((Xp(x) - Xp(Y)) ^ 2 + (Yp(x) - Yp(Y)) ^ 2) < (2 * Rad) ^ 2 Then
                    Ang = ATan2(Yp(x) - Yp(Y), Xp(x) - Xp(Y))
                    Cx = (Xp(x) + Xp(Y)) / 2
                    Cy = (Yp(x) + Yp(Y)) / 2
                    Xp(x) = Cx + (Rad * Cos(Ang))
                    Yp(x) = Cy - (Rad * Sin(Ang))
                    Xp(Y) = Cx - (Rad * Cos(Ang))
                    Yp(Y) = Cy + (Rad * Sin(Ang))
                End If
            End If
        Next Y
    Next x
    Draw
End Sub

Private Sub Draw()
    Cls
    For x = 0 To Balls
        Circle (Xp(x), Yp(x)), Rad, Clr(x)
    Next x
End Sub
