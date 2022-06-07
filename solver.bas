Attribute VB_Name = "Module1"
Option Explicit

Dim t(81) As Integer
Dim mem(81) As Integer

Sub main()

End Sub
    Dim i As Integer, k As Integer
    Dim h As Integer
    Dim flg As Integer
    ' flg = 1:ç∂Ç…êiÇﬁ 2:âEÇ…êiÇﬁ
    Dim chk As Boolean
    
    
    Application.ScreenUpdating = False
    Set blk = Range("A1:I9")
    For i = 1 To 81
        t(i) = blk(i).Value
        If t(i) = 0 Then
            mem(i) = 0
        Else
            mem(i) = 1
        End If
    Next i
    
    i = 1
    
    flg = 2
    Do While i <= 81
        If mem(i) = 0 Then
            If flg = 1 Then
                h = t(i)
                k = 1
                Do While (h + k) <= 9
                    chk = check(i, h + k)
                    If chk Then
                        t(i) = h + k
                        flg = 2
                        Exit Do
                    End If
                    k = k + 1
                Loop
            Else
                flg = 1
                For h = 1 To 9
                    chk = check(i, h)
                    If chk Then
                        t(i) = h
                        flg = 2
                        Exit For
                    End If
                Next h
            End If
        End If
        
        If flg = 2 Then
            i = i + 1
        Else
            If mem(i) = 0 Then
                t(i) = 0
            End If
            i = i - 1
        End If
    Loop
    
    For i = 1 To 81
        If mem(i) = 0 Then
            With blk(i)
                .Value = t(i)
                .Font.ColorIndex = 3
            End With
        End If
    Next i

End Sub
Function src_block(index As Integer, clm As Integer) As Integer
    src_block = (((index - 1) \ 27) * 27 + 1) + (((clm - 1) \ 3) * 3)
End Function

Function src_row(index As Integer) As Integer
    src_row = (((index - 1) \ 9) * 9) + 1
End Function

Function src_colmun(index As Integer) As Integer
    src_colmun = ((index - 1) Mod 9) + 1
End Function

Function check(index As Integer, h As Integer) As Boolean
    Dim i As Integer
    
    Dim row_num As Integer
    row_num = src_row(index)
    
    For i = 0 To 8
        If t(row_num + i) = h Then
            check = False
            Exit Function
        End If
    Next i
    
    Dim column_num As Integer
    column_num = src_colmun(index)
    
    For i = 0 To 72 Step 9
        If t(column_num + i) = h Then
            check = False
            Exit Function
        End If
    Next i
    
    Dim block_num As Integer, j As Integer
    block_num = src_block(index, column_num)
    
    For j = 0 To 18 Step 9
        For i = 0 To 2
            If t(block_num + i + j) = h Then
                check = False
                Exit Function
            End If
        Next i
    Next j
    
    check = True

End Function

Sub delete_red()
    Dim i As Integer
    Dim blk As Object
    
    Set blk = Worksheets("êîì∆").Range("A1:I9")
    Application.ScreenUpdating = False
    For i = 1 To 81
        With blk(i)
            If .Font.ColorIndex = 3 Then
                .ClearContents
                .Font.ColorIndex = 1
            End If
        End With
    Next i
End Sub





