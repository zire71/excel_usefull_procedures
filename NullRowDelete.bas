Public Sub nullRowDelete() 'Процедура удаления пустых строк
    
    'Процедура удалеяет пустые строки в выделенном диапазоне ячеек

    Dim ir As Long, ic%, startRow As Long, startColumn%, i As Long, j%, tempArr() As String, inull%, nullSearch$, rowsArr() As Long, irArr%
        
    startRow = Selection.Row
    startColumn = Selection.Column
    ir = Selection.Rows.Count
    ic = Selection.Columns.Count
    'array
    irArr = 1
    For i = 0 To ir - 1
        
        inull = 1
        For j = 0 To ic - 1
            
            ReDim Preserve tempArr(0 To j)
            
            tempArr(j) = ActiveSheet.Cells(startRow + i, startColumn + j).Value
            
                    
        Next j
        
        nullSearch = Join(tempArr, "")
        
        If nullSearch Like "" Then
            
            ReDim Preserve rowsArr(1 To irArr)
            rowsArr(irArr) = startRow + i
            irArr = irArr + 1
            
        End If
    Next i
    
    For i = UBound(rowsArr) To 1 Step -1
        'MsgBox rowsArr(i)
        ActiveSheet.Range(ActiveSheet.Cells(rowsArr(i), startColumn).Address & ":" & _
                    ActiveSheet.Cells(rowsArr(i), startColumn + ic - 1).Address).Delete Shift:=xlUp
    Next i
    
End Sub
