Sub ГенераторЧисел()
    Dim arr(1 To 10, 1 To 10) As Integer
    Const ROWS_COUNT As Integer = 10
    Const COLS_COUNT As Integer = 10
    Const MAX_VALUE As Integer = 100
    Dim rand As Integer
    Dim row As Integer
    Dim column As Integer
    Dim min As Integer
    Dim max As Integer
    Dim cellMin As String
    Dim cellMax As String
    
    min = MAX_VALUE
    Randomize
    
    For row = 1 To ROWS_COUNT
        For column = 1 To COLS_COUNT
            rand = Int((100 * Rnd) + 1)
            arr(row, column) = rand
            If rand > max Then
            max = rand
            cellMax = Cells(row, column).Address
            End If
            If rand < min Then
            min = rand
            cellMin = Cells(row, column).Address
            End If
        Next column
    Next row
    
    Range("A1:J10").Font.Color = RGB(255, 255, 255)
    Range("A1:J10").Value = arr
    
        For row = 1 To ROWS_COUNT
            For column = 1 To COLS_COUNT
        If row Mod 2 = 0 And column Mod 2 = 0 Then
        Cells(row, column).Interior.Color = RGB(0, Интенсивность(arr(row, column), MAX_VALUE), 0)
        ElseIf row Mod 2 = 1 And column Mod 2 = 1 Then
        Cells(row, column).Interior.Color = RGB(0, 0, Интенсивность(arr(row, column), MAX_VALUE))
        Else:
        Cells(row, column).Interior.Color = RGB(Интенсивность(arr(row, column), MAX_VALUE), 0, 0)
        End If

            Next column
        Next row
        
    Cells(1, 11).Value = "MIN:"
    Cells(2, 11).Value = min
    Cells(3, 11).Value = cellMin
    
    Cells(1, 12).Value = "MAX:"
    Cells(2, 12).Value = max
    Cells(3, 12).Value = cellMax
End Sub
    Function Интенсивность(значение As Integer, максимум As Integer)
        Интенсивность = 255 * значение / максимум
    End Function
