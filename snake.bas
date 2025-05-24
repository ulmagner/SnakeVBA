Attribute VB_Name = "Module1"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Const VK_UP = &H26
Const VK_DOWN = &H28
Const VK_LEFT = &H25
Const VK_RIGHT = &H27

Dim X As Integer, Y As Integer
Dim Dx As Integer, Dy As Integer
Sub Snake()
    Dim Win As Boolean
    Dim Rdx As Integer
    Dim Rdy As Integer

    X = 12
    Y = 18
    Dx = 0
    Dy = -1
    Win = False
    Do While Not Intersect(Cells(X, Y), Range("C3:Z20")) Is Nothing
        If GetAsyncKeyState(VK_UP) < 0 Then
            UpDirection
        ElseIf GetAsyncKeyState(VK_DOWN) < 0 Then
            DownDirection
        ElseIf GetAsyncKeyState(VK_LEFT) < 0 Then
            LeftDirection
        ElseIf GetAsyncKeyState(VK_RIGHT) < 0 Then
            RightDirection
        End If
        If Cells(X + Dx, Y + Dy).Value = Cells(2, 30) Then
            Cells(6, 31).Value = Cells(6, 31).Value + 1
            Rdx = WorksheetFunction.RandBetween(2, 20)
            Rdy = WorksheetFunction.RandBetween(2, 27)
            Cells(Rdx, Rdy).Value = Cells(2, 30)
        End If
        If Cells(6, 31).Value = 10 Then
            Win = True
            Exit Do
        End If
        Cells(X + Dx, Y + Dy).Value = Cells(X, Y).Value
        Cells(X, Y).ClearContents
        X = X + Dx
        Y = Y + Dy
        DoEvents
        Sleep 300
    Loop

    If Win = False Then
        MsgBox "You Lose."
    Else
        MsgBox "Congratulations, You Win!!!"
    End If
End Sub

Sub UpDirection()
    If Dx = 0 Then
        Dx = -1
        Dy = 0
    End If
End Sub

Sub DownDirection()
    If Dx = 0 Then
        Dx = 1
        Dy = 0
    End If
End Sub

Sub LeftDirection()
    If Dy = 0 Then
        Dx = 0
        Dy = -1
    End If
End Sub

Sub RightDirection()
    If Dy = 0 Then
        Dx = 0
        Dy = 1
    End If
End Sub

Sub Reset()
    Dim Rdx As Integer
    Dim Rdy As Integer

    Range("B2:AA21").ClearContents
    Range("AE6") = 0
    Range("AE16") = 1
    Rdx = WorksheetFunction.RandBetween(3, 20)
    Rdy = WorksheetFunction.RandBetween(3, 26)
    Cells(Rdx, Rdy).Value = Cells(2, 30)
    Rdx = WorksheetFunction.RandBetween(3, 20)
    Rdy = WorksheetFunction.RandBetween(3, 26)
    Cells(Rdx, Rdy).Value = Cells(2, 30)
    Rdx = WorksheetFunction.RandBetween(3, 20)
    Rdy = WorksheetFunction.RandBetween(3, 26)
    Cells(Rdx, Rdy).Value = Cells(2, 30)
    Range("R12") = Range("AD3")
End Sub
