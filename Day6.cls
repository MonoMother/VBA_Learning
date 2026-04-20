Sub УдалитьПустыеСтрокиИПодсветитьПоследнюю()
    Dim LastRow As Long
    Dim i As Long
    
    Cells.Interior.ColorIndex = xlNone
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = LastRow To 1 Step -1
        If Trim(Cells(i, 1).Value) = "" Then
            Rows(i).Delete
        End If
    Next i
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(LastRow, 1).Interior.Color = RGB(0, 255, 0)
    
    MsgBox "Последняя строка: " & LastRow
End Sub
