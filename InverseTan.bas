Attribute VB_Name = "InverseTan"
Public Const pi As Double = 3.14159265358978

Public Function ATan2(Y, x)
    ATan2 = vbNull
    If x = 0 And Y = 0 Then
        Exit Function
    ElseIf x = 0 And Y < 0 Then
        ATan2 = pi / 2
    ElseIf x < 0 Then
        ATan2 = pi - Atn(Y / x)
    ElseIf x = 0 And Y > 0 Then
        ATan2 = -pi / 2
    ElseIf x > 0 Then
        ATan2 = 2 * pi - Atn(Y / x)
    End If
End Function

Public Function Rad(angleinput As Double) As Double
    Rad = angleinput * (pi / 180)
End Function

Public Function Angl(radinput As Double) As Double
    Angl = radinput * (180 / pi)
End Function

Public Sub Bnc(ByRef Ang1 As Double, Wall As Double)
    Ang1 = Wall + Wall - Ang1
End Sub

Public Function DstSq(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Single
    DstSq = ((x1 - x2) * (x1 - x2)) + ((y1 - y2) * (y1 - y2))
End Function

Public Function Dst(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Single
    Dst = Sqr(DstSq(x1, y1, x2, y2))
End Function


