Function LastRow(ws As Worksheet, col As Long) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function


Sub ВыделитьПоследнююСтроку()
    Dim ws As Worksheet
    Dim последняяСтрока As Long
    
    Set ws = ActiveSheet
    последняяСтрока = LastRow(ws, 1)
    If последняяСтрока = 1 And ws.Cells(1, 1).Value = "" Then
    MsgBox "Таблица пуста!"
    End If
    
    ws.Range("A:Z").Font.Bold = False
    
    ws.Range("A" & последняяСтрока & ":Z" & последняяСтрока).Font.Bold = True
    
    MsgBox "строка " & последняяСтрока & " выделена жирным "
End Sub
