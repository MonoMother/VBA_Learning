Sub Задачи_Расширенная()
    Dim lastRow As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim просрочено As Long
    Dim готово As Long
    
    Set ws = Worksheets("Лист2")
    ws.Range("A:Z").Interior.ColorIndex = xlNone
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Сбор статистики
    просрочено = 0
    готово = 0
    
    For i = lastRow To 1 Step -1
        ' Пропуск заголовка
        If i = 1 Then GoTo пропустить
        
        ' Защита от пустых и неверных дат
        If Not IsDate(ws.Cells(i, 1).Value) Then GoTo пропустить
        
        ' Подсчёт статистики
        If ws.Cells(i, 2).Value = "Готово" Then
            готово = готово + 1
        End If
        
        ' Основное условие
        If ws.Cells(i, 1).Value < Date And ws.Cells(i, 2).Value <> "Готово" Then
            ws.Rows(i).Interior.Color = RGB(255, 0, 0)
            просрочено = просрочено + 1
        End If
        
пропустить:
    Next i
    
    ' Итоговое сообщение с данными
    MsgBox "Готово!" & vbNewLine & _
           "Просроченных задач: " & просрочено & vbNewLine & _
           "Выполненных задач: " & готово
    
    ' Очистка цветов горячей клавишей 
    MsgBox "Чтобы очистить цвета, нажмите Ctrl+Shift+R"
End Sub

' Отдельный макрос для очистки цветов
Sub ОчиститьЦвета()
    Dim ws As Worksheet
    Set ws = Worksheets("Лист2")
    ws.Range("A:Z").Interior.ColorIndex = xlNone
    MsgBox "Цвета очищены"
End Sub
