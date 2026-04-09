Sub СчётчикСлов()
    Dim dict As Object
    Dim arr() As String
    Dim i As Integer
    Dim слово As Variant
    Dim строка As Integer
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    arr = Split(Range("A1").Value, " ")
    
    For i = LBound(arr) To UBound(arr)
        
        If dict.Exists(arr(i)) Then
            dict(arr(i)) = dict(arr(i)) + 1
        Else
            dict.Add arr(i), 1
        End If
        
    Next i
    
    строка = 1
    
    For Each слово In dict.Keys
        Cells(строка, 3).Value = слово      ' столбец C
        Cells(строка, 4).Value = dict(слово) ' столбец D
        строка = строка + 1
    Next слово
    
    MsgBox "Найдено уникальных слов: " & dict.Count
End Sub
