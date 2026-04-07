Sub УмнаяОчисткаТаблицы()
    ' =====================================================
    ' ОЧИЩАЕТ ТАБЛИЦУ ОТ МУСОРА И ПОДГОТАВЛИВАЕТ К АНАЛИЗУ
    ' =====================================================
    
    Dim i As Long
    Dim LastRow As Long
    Dim ЧисткаВыполнена As Boolean
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = LastRow To 1 Step -1
        Dim Значение As String
        Значение = Trim(Cells(i, 1).Value)  
        
        If Значение = "" Then
            Rows(i).Delete
            ЧисткаВыполнена = True
            
        ElseIf Значение = "---" Or Значение = "***" Or Значение = "ИТОГО" Then
            Rows(i).Delete
            ЧисткаВыполнена = True
            
        ElseIf LCase(Left(Значение, 5)) = "итого" Or LCase(Left(Значение, 5)) = "всего" Then
            Rows(i).Delete
            ЧисткаВыполнена = True
        End If
    Next i
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        If IsNumeric(Cells(i, 2).Value) Then
        ElseIf IsNumeric(Trim(Cells(i, 2).Value)) Then
            Cells(i, 2).Value = Val(Trim(Cells(i, 2).Value))
        Else
            Cells(i, 2).Interior.Color = RGB(255, 200, 200)
        End If
    Next i
    
    For i = 2 To LastRow
        Dim Цена As Variant
        Dim ОчищеннаяЦена As String
        Dim j As Integer
        
        Цена = Cells(i, 3).Value
        ОчищеннаяЦена = ""

        If Цена <> "" And Not IsNull(Цена) Then

            Dim СтрокаЦены As String
            СтрокаЦены = CStr(Цена)
            
            For j = 1 To Len(СтрокаЦены)
                Dim Символ As String
                Символ = Mid(СтрокаЦены, j, 1)
                

                If (Символ >= "0" And Символ <= "9") Or Символ = "." Or Символ = "," Then
                    ОчищеннаяЦена = ОчищеннаяЦена & Символ
                End If
            Next j
            

            ОчищеннаяЦена = Replace(ОчищеннаяЦена, ",", ".")
            

            If IsNumeric(ОчищеннаяЦена) And ОчищеннаяЦена <> "" Then
                Cells(i, 3).Value = CDbl(ОчищеннаяЦена)
                Cells(i, 3).Interior.ColorIndex = xlNone
            Else

                Cells(i, 3).Interior.Color = RGB(255, 200, 200)
            End If
        End If
    Next i
    
    For i = LastRow To 2 Step -1
        If IsNumeric(Cells(i, 2).Value) Then
            If Cells(i, 2).Value <= 0 Then
                Rows(i).Delete
                ЧисткаВыполнена = True
            End If
        End If
    Next i
    

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    If ЧисткаВыполнена Then
        MsgBox "Очистка завершена!" & vbNewLine & _
               "Осталось строк: " & LastRow - 1 & vbNewLine & _
               "Проблемные ячейки подсвечены красным."
    Else
        MsgBox "Таблица уже чистая. Ничего не удалено."
    End If
    
End Sub
