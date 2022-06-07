Attribute VB_Name = "Module2"
Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As LongPtr) As Long
Dim time As Integer, tm1 As Integer
Const tm2 As Integer = 30
Public mem(3, 1) As Integer
Public score As Long

Sub main()
    Dim cnt As Integer, next_num As Integer
    Randomize
    
    Call clear
    next_num = 0
    score = 0
    tm1 = Cells(2, "Y").Value
    Do While True
        time = tm1
        next_num = BLOCK_choise(next_num)
        Cells(2, 18).Value = score
        If Cells(3, 8).Interior.ColorIndex <> -4142 Or Cells(3, 9).Interior.ColorIndex <> -4142 Or Cells(3, 10).Interior.ColorIndex <> -4142 Then
            MsgBox "GAMEOVER"
            End
        End If
    Loop

End Sub

Function BLOCK_choise(next_num As Integer) As Integer
    Dim num As Integer
    If next_num = 0 Then
        num = Int((7 - 1 + 1) * Rnd + 1)
        next_num = Int((7 - 1 + 1) * Rnd + 1)
    Else
        num = next_num
        next_num = Int((7 - 1 + 1) * Rnd + 1)
    End If
    Range("R6:V9").Interior.ColorIndex = 0
    
    Select Case next_num
        Case 1
            Union(Cells(8, "S"), Cells(7, "S"), Cells(8, "T"), Cells(7, "T")).Interior.ColorIndex = 10
        Case 2
            Union(Cells(7, "T"), Cells(7, "U"), Cells(7, "R"), Cells(7, "S")).Interior.ColorIndex = 8
        Case 3
            Union(Cells(7, "T"), Cells(8, "T"), Cells(8, "S"), Cells(8, "U")).Interior.ColorIndex = 29
        Case 4
            Union(Cells(8, "S"), Cells(8, "T"), Cells(8, "U"), Cells(7, "U")).Interior.ColorIndex = 45
        Case 5
            Union(Cells(7, "S"), Cells(8, "T"), Cells(8, "U"), Cells(8, "S")).Interior.ColorIndex = 5
        Case 6
            Union(Cells(8, "S"), Cells(8, "T"), Cells(7, "T"), Cells(7, "U")).Interior.ColorIndex = 4
        Case 7
            Union(Cells(7, "S"), Cells(7, "T"), Cells(8, "T"), Cells(8, "U")).Interior.ColorIndex = 3
    End Select
    
    Call BLOCK_move(num)
    
    BLOCK_choise = next_num
End Function

Sub BLOCK_move(num As Integer)
    Dim x As Integer, y As Integer, state As Integer, s_y As Integer
    Const start_x As Integer = 4
    Const start_y As Integer = 9
    
    Dim mn As Object
    
    Select Case num
        Case 1
            
            Set mn = New Mino1
        Case 2
            
            Set mn = New Mino2
        Case 3
            
            Set mn = New Mino3
        Case 4
            
            Set mn = New Mino4
        Case 5
            
            Set mn = New Mino5
        Case 6
            
            Set mn = New Mino6
        Case 7
            
            Set mn = New Mino7
    End Select
    
    x = start_x
    y = start_y
    
    state = 0
    
    Do While True
        time = tm1
        
        If mn.Vmove() Then
            time = tm2
        End If
        s_y = y
        mn.rotate state, x, y
        mn.make x, y, state
        Sleep time
        
        If mn.Bcheck(x, y, state) Then
            y = y + mn.Hmove(x, y, state)
            If y <> s_y Then
                GoTo continue
            End If
            score = mn.bingo(x, state) * 100 + score
            Sleep time
            
            Exit Do
        End If
        y = y + mn.Hmove(x, y, state)
continue:
        mn.delete
        If s_y = y Then '  suiheiidou
            x = x + 1
        End If
    Loop
    Set mn = Nothing
End Sub


Sub clear()

    Range("E3:N23").Interior.ColorIndex = 0
    Range("R6:V9").Interior.ColorIndex = 0
    Union(Range("D3:D24"), Range("E24:N24"), Range("O3:O24")).Interior.ColorIndex = 16
    Cells(2, 18).Value = 0
End Sub

