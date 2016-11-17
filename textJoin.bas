Public Sub textJoin()
    'Процедура объединяет текст из выделенных ячеек в одном столбце в первую ячейку
    'при этом очищаяя все ячейки ниже первой в выделении
    
    Dim tempStr$, i As Long, tr As Long, tc%
    
    tr = Selection.Row
    tc = Selection.Column
    
    tempStr = ActiveSheet.Cells(tr, tc)
    
    If Selection.Rows.Count < 2 Then
        MsgBox "Вы выбрали слишком маленький диапазон. Всего - " & Selection.Rows.Count & _
         _ " строка. Выполнение процедуры невозможно. Выберите больше строк."
        End Sub
    End If
    
    For i = 1 To Selection.Rows.Count - 1
        tempStr = tempStr & " " & ActiveSheet.Cells(tr + i, tc)
        ActiveSheet.Cells(tr + i, tc) = Null
    Next i
    
    ActiveSheet.Cells(tr, tc) = tempStr
    
End Sub
