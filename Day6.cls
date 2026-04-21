Sub Dictionary()
    Dim dict As Object
    Dim product As String
    Dim sales As Double
    Dim ws As Worksheet
    
    Set ws = Worksheets("Sales")
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        product = ws.Cells(i, 1).Value
        sales = ws.Cells(i, 2).Value
            If dict.Exists(product) Then
                dict(product) = dict(product) + sales
            Else
                dict.Add product, sales
            End If
    Next i
    
    Dim wsSum As Worksheet
    On Error Resume Next
    
    Set wsSum = Worksheets("Summary")
        If wsSum Is Nothing Then
        Set wsSum = Worksheets.Add
        wsSum.Name = "Summary"
        Else
        wsSum.Cells.Clear
        End If
    
    On Error GoTo 0
    
    wsSum.Cells(1, 1).Value = "Продукт"
    wsSum.Cells(1, 2).Value = "Сумма"
    
    Dim rows As Integer
    rows = 2
    
    Dim key As Variant
    
    For Each key In dict.Keys
        wsSum.Cells(rows, 1).Value = key
        wsSum.Cells(rows, 2).Value = dict(key)
        rows = rows + 1
    Next key
    MsgBox "Готово! Обработано " & dict.Count & "продуктов"
        
End Sub
