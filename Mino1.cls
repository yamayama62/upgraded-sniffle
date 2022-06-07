VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mino1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub make(x, y, state)
    Union(Cells(x, y), Cells(x, y + 1), Cells(x - 1, y), Cells(x - 1, y + 1)).Interior.ColorIndex = 10
    mem(0, 0) = x
    mem(0, 1) = y
    mem(1, 0) = x
    mem(1, 1) = y + 1
    mem(2, 0) = x - 1
    mem(2, 1) = y
    mem(3, 0) = x - 1
    mem(3, 1) = y + 1
End Sub

Public Function Hmove(ByVal x, ByVal y As Integer, ByVal state) As Long

    If GetAsyncKeyState(vbKeyLeft) <> 0 Then
        If Cells(x, y - 1).Interior.ColorIndex = -4142 And Cells(x - 1, y - 1).Interior.ColorIndex = -4142 Then
            Hmove = -1
        End If
    ElseIf GetAsyncKeyState(vbKeyRight) <> 0 Then
        If Cells(x, y + 2).Interior.ColorIndex = -4142 And Cells(x - 1, y + 2).Interior.ColorIndex = -4142 Then
            Hmove = 1
        End If
    End If
End Function

Public Function Vmove() As Boolean
    If GetAsyncKeyState(vbKeyDown) <> 0 Then
        Vmove = True
    End If
End Function
Public Sub rotate(ByRef state, ByVal x, ByVal y)
    'pass
End Sub
Public Sub delete()
    Dim i As Integer
    
    For i = 0 To 3
        Cells(mem(i, 0), mem(i, 1)).Interior.ColorIndex = 0
    Next i
End Sub

Public Function Bcheck(x As Integer, y As Integer, state As Integer) As Boolean

    If Cells(x + 1, y).Interior.ColorIndex <> -4142 Or Cells(x + 1, y + 1).Interior.ColorIndex <> -4142 Then
        Bcheck = True
    End If
End Function

Public Function bingo(x As Integer, state As Integer) As Long
    Dim i As Integer, j As Integer, cnt As Integer
    Dim flg As Boolean
    Dim list(2) As Integer
    Dim s As Long
    
    s = 0
    cnt = 2
    For i = -1 To 0
        flg = True
        For j = 5 To 14
            If Cells(x + i, j).Interior.ColorIndex = -4142 Then
                flg = False
                Exit For
            End If
        Next j
        If flg = True Then
            If s = 0 Then
                s = 1
            Else
                s = s * 2
            End If
            list(i + 1) = x + i
        End If
    Next i
    
    For i = 0 To (cnt - 1)
        If list(i) <> 0 Then
            Range(Cells(3, 5), Cells(list(i) - 1, 14)).Cut
            Cells(4, 5).Select
            ActiveSheet.Paste
            Cells(1, 1).Select
            Application.CutCopyMode = False
        End If
    Next i
    
    bingo = s
End Function

