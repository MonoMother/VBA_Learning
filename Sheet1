Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 2 Then
        

        If IsNumeric(Target.Value) Then

            If Target.Value = 10 Then
                Target.Offset(0, 1).Value = "СРОЧНО!!!"
                Target.Offset(0, 1).Interior.Color = RGB(255, 0, 0)
                Target.Offset(0, 2).Value = "Высокий"
            ElseIf Target.Value >= 1 And Target.Value < 7 Then
                Target.Offset(0, 1).Value = "Обычное"
                Target.Offset(0, 1).Interior.Color = RGB(150, 150, 0)
                Target.Offset(0, 2).Value = "Низкий"
            ElseIf Target.Value >= 7 And Target.Value < 10 Then
                Target.Offset(0, 1).Value = "Важно"
                Target.Offset(0, 1).Interior.Color = RGB(255, 255, 0)
                Target.Offset(0, 2).Value = "Средний"
            Else
                ' Число меньше 1
                Target.Value = ""
                MsgBox ("Ошибка: это не число")
            End If
            
        End If
        End If
        
End Sub
