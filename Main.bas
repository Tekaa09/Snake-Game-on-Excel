Attribute VB_Name = "Main"
Public rinc As Integer, cinc As Integer
Dim r() As Integer, c() As Integer
Public Moving As Boolean
Public apples As Integer
Public nextTick As Date

Sub StartGame()
Range("B2:Q34").Interior.Color = vbBlack
ReDim r(2)
ReDim c(2)
r(0) = 20: r(1) = 21: r(2) = 22
c(0) = 10: c(1) = 10: c(2) = 10
apples = 0
rinc = -1
cinc = 0
Moving = True
GameLoop1
ShowSnake
bindKeys
AddApple
End Sub

Sub ShowSnake()
Application.ScreenUpdating = False
For i = UBound(r) To 1 Step -1
  If i Mod 2 = 0 Then
     Cells(r(i), c(i)).Interior.Color = vbRed
  Else
     Cells(r(i), c(i)).Interior.Color = vbGreen
  End If
Next i
Cells(r(0), c(0)).Interior.Color = vbRed
Application.ScreenUpdating = True
End Sub
Sub MoveSnake()
    Dim i As Integer, tail As Integer

    If rinc <> 0 Or cinc <> 0 Then
        tail = UBound(r)
        Cells(r(tail), c(tail)).Interior.Color = vbBlack

        For i = tail To 1 Step -1
            r(i) = r(i - 1)
            c(i) = c(i - 1)
        Next i

        r(0) = r(0) + rinc
        c(0) = c(0) + cinc

        If r(0) < 2 Then r(0) = 34
        If r(0) > 34 Then r(0) = 2
        If c(0) < 2 Then c(0) = 17
        If c(0) > 17 Then c(0) = 2
                
        For i = 1 To UBound(r)
           If r(0) = r(i) And c(0) = c(i) Then
               Moving = False
               MsgBox "Game Over!"
               Exit Sub
            End If
        Next i
        If Cells(r(0), c(0)).Interior.Color = vbBlue Then
               Moving = False
               MsgBox "Game Over!"
               Exit Sub
        End If
        If Cells(r(0), c(0)).Interior.Color = vbYellow Then
          apples = apples + 1
          ReDim Preserve r(UBound(r) + 1)
          ReDim Preserve c(UBound(c) + 1)
          r(UBound(r)) = r(UBound(r) - 1)
          c(UBound(c)) = c(UBound(c) - 1)
          AddApple
          AddBoom
        End If
       ShowSnake
       UpdateAppleButton
   End If
End Sub

Sub AddApple()
    Dim found As Boolean
    Do
        found = False
        appleR = Int((34 - 2 + 1) * Rnd + 2)
        appleC = Int((17 - 2 + 1) * Rnd + 2)
        Dim i As Integer
        For i = 0 To UBound(r)
            If r(i) = appleR And c(i) = appleC Then
                found = True
                Exit For
            End If
        Next i
    Loop While found
    Cells(appleR, appleC).Interior.Color = vbYellow
End Sub
Sub UpdateAppleButton()
    Dim btn As Button
    Set btn = ActiveSheet.Buttons("Button 8")
    btn.Text = "Points: " & apples * 20
End Sub

Sub GameLoop1()
    If Not Moving Then Exit Sub
    MoveSnake
    nextTick = DateAdd("s", 1, Now)
    Application.OnTime nextTick, "GameLoop1"
End Sub


Sub AddBoom()
    Dim found1 As Boolean
    Do
        found1 = False
        BoomR = Int((34 - 2 + 1) * Rnd + 2)
        Boomc = Int((17 - 2 + 1) * Rnd + 2)
        Dim i As Integer
        For i = 0 To UBound(r)
            If r(i) = BoomR And c(i) = Boomc Then
                found1 = True
                Exit For
            End If
        Next i
    Loop While found
    Cells(BoomR, Boomc).Interior.Color = vbBlue
End Sub
