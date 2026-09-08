Attribute VB_Name = "Enemy"
Public rinc1 As Integer, cinc1 As Integer
Dim r1() As Integer, c1() As Integer
Public Moving As Boolean

Sub EnemyStart()
Dim ri As Integer
Randomize
ri = Int((34 - 2 + 1) * Rnd) + 2
ReDim r1(2)
ReDim c1(2)
r1(0) = ri + 2: r1(1) = ri + 1: r1(2) = ri
c1(0) = 2: c1(1) = 2: c1(2) = 2
apples = 0
rinc1 = 0
cinc1 = 0
Moving = True
gameRunning = True
GameLoop
ShowSnake
End Sub
Sub ShowSnake()
ClearSnake
For i = UBound(r1) To 1 Step -1
  If i Mod 2 = 0 Then
     Cells(r1(i), c1(i)).Interior.Color = vbBlue
  Else
     Cells(r1(i), c1(i)).Interior.Color = vbCyan
  End If
Next i
Cells(r1(0), c1(0)).Interior.Color = vbBlue
End Sub
Sub MoveSnake1()
    Dim i As Integer, tail As Integer
    If rinc1 <> 0 Or cinc1 <> 0 Then
        tail = UBound(r1)
        Cells(r1(tail), c1(tail)).Interior.Color = vbBlack
        For i = tail To 1 Step -1
            r1(i) = r1(i - 1)
            c1(i) = c1(i - 1)
        Next i
        r1(0) = r1(0) + rinc1
        c1(0) = c1(0) + cinc1
        
        If r1(0) < 2 Then r1(0) = 34
        If r1(0) > 34 Then r1(0) = 2
        If c1(0) < 2 Then c1(0) = 12
        If c1(0) > 12 Then c1(0) = 2

       ShowSnake
   End If
End Sub

Sub GameLoop()
    If Not Moving Then Exit Sub
    MoveSnake1
    nextTick = DateAdd("s", 0.6, Now)
    Application.OnTime nextTick, "GameLoop"
End Sub

Sub ClearSnake()
    Dim i As Integer
    For i = 0 To UBound(r1)
        Cells(r1(i), c1(i)).Interior.Color = vbBlack
    Next i
End Sub
