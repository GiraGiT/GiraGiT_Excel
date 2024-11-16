Sub CombineRanges()
    Dim choice As Variant
    Dim rangeValues As Range
    Dim criteria As Range
    Dim rangeSum As Range
    Dim rangeValues2 As Range
    Dim criteria2 As Range
    Dim countRange As Range
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    choice = InputBox("Выберите действие: введите '1' для суммирования, '2' - для суммирования по двум критериям, '3' - для подсчета уникальных значений, '4' - для подсчета значений по цвету, '5' - для подсчета всех значений по цвету")
    
    If choice = "" Then Exit Sub
    
    Set rangeValues = Application.InputBox("Выберите диапазон для поиска значения и цвета", Type:=8)
    If rangeValues Is Nothing Then Exit Sub
    
    Set criteria = Application.InputBox("Выберите ячейку с значением и цветом для поиска", Type:=8)
    If criteria Is Nothing Then Exit Sub
    
    Select Case choice
        Case "1"
            Set rangeSum = Application.InputBox("Выберите диапазон для суммирования", Type:=8)
            If rangeSum Is Nothing Then Exit Sub
            ActiveCell.Formula = "=SumIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & rangeSum.Columns(1).Address & ")"
            
        Case "2"
            Set rangeValues2 = Application.InputBox("Выберите второй диапазон для поиска значения и цвета", Type:=8)
            If rangeValues2 Is Nothing Then Exit Sub
            
            Set criteria2 = Application.InputBox("Выберите вторую ячейку с значением и цветом для поиска", Type:=8)
            If criteria2 Is Nothing Then Exit Sub
            
            Set rangeSum = Application.InputBox("Выберите диапазон для суммирования", Type:=8)
            If rangeSum Is Nothing Then Exit Sub
            
            ActiveCell.Formula = "=SumIfColorAndValue2(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & rangeSum.Columns(1).Address & ", " & rangeValues2.Columns(1).Address & ", " & criteria2.Address & ")"
            
        Case "3"
            Set countRange = Application.InputBox("Выберите диапазон для подсчета уникальных значений", Type:=8)
            If countRange Is Nothing Then Exit Sub
            ActiveCell.Formula = "=CountUniqueIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & countRange.Columns(1).Address & ")"
            
        Case "4"
            Set rangeValues2 = Application.InputBox("Выберите второй диапазон для поиска значения и цвета", Type:=8)
            If rangeValues2 Is Nothing Then Exit Sub
            
            Set criteria2 = Application.InputBox("Выберите вторую ячейку с значением и цветом для поиска", Type:=8)
            If criteria2 Is Nothing Then Exit Sub
            
            ActiveCell.Formula = "=CountIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & rangeValues2.Columns(1).Address & ", " & criteria2.Address & ")"
            
        Case "5"
            Set countRange = Application.InputBox("Выберите диапазон для подсчета всех значений", Type:=8)
            If countRange Is Nothing Then Exit Sub
            ActiveCell.Formula = "=CountAllIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & countRange.Columns(1).Address & ")"
    End Select
    
    Exit Sub

ErrorHandler:
    ' Формируем подробное сообщение об ошибке
    errorMsg = "Произошла ошибка в процедуре CombineRanges:" & vbNewLine & _
               "Номер ошибки: " & Err.Number & vbNewLine & _
               "Описание: " & Err.Description & vbNewLine & _
               "Активный лист: " & ActiveSheet.Name & vbNewLine & _
               "Активная ячейка: " & Selection.Address
    
    ' Показываем сообщение об ошибке
    MsgBox errorMsg, vbCritical, "Ошибка в CombineRanges"
End Sub

Function SumIfColorAndValue(rangeValues As Range, criteria As Range, rangeSum As Range) As Double
    Dim cell As Range
    Dim i As Long
    Dim color As Long
    Dim errorMsg As String
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = rangeValues.Worksheet
    color = criteria.Interior.color
    
    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).value = criteria.value And rangeValues.Cells(i).Interior.color = color Then
            If IsNumeric(rangeSum.Cells(i).value) Then
                SumIfColorAndValue = SumIfColorAndValue + rangeSum.Cells(i).value
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    errorMsg = "Произошла ошибка в функции SumIfColorAndValue:" & vbNewLine & _
               "Номер ошибки: " & Err.Number & vbNewLine & _
               "Описание: " & Err.Description & vbNewLine & _
               "Лист: " & ws.Name & vbNewLine & _
               "Ячейка: " & rangeValues.Cells(i).Address
    
    MsgBox errorMsg, vbCritical, "Ошибка в SumIfColorAndValue"
    SumIfColorAndValue = 0 ' Возвращаем 0 в случае ошибки
End Function

Function SumIfColorAndValue2(rangeValues1 As Range, criteria1 As Range, rangeSum As Range, rangeValues2 As Range, criteria2 As Range) As Double
    Dim i As Long
    Dim color1 As Long
    Dim cellValue1 As Variant, cellValue2 As Variant
    Dim sumValue As Variant
    Dim errorMsg As String
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = rangeValues1.Worksheet
    color1 = criteria1.Interior.color
    
    For i = 1 To rangeValues1.Cells.count
        cellValue1 = rangeValues1.Cells(i).value
        cellValue2 = rangeValues2.Cells(i).value
        If cellValue1 = criteria1.value And rangeValues1.Cells(i).Interior.color = color1 And cellValue2 = criteria2.value Then
            sumValue = rangeSum.Cells(i).value
            If IsNumeric(sumValue) Then
                SumIfColorAndValue2 = SumIfColorAndValue2 + sumValue
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    errorMsg = "Произошла ошибка в функции SumIfColorAndValue2:" & vbNewLine & _
               "Номер ошибки: " & Err.Number & vbNewLine & _
               "Описание: " & Err.Description & vbNewLine & _
               "Лист: " & ws.Name & vbNewLine & _
               "Ячейка: " & rangeValues1.Cells(i).Address
    
    MsgBox errorMsg, vbCritical, "Ошибка в SumIfColorAndValue2"
    SumIfColorAndValue2 = 0 ' Возвращаем 0 в случае ошибки
End Function

Function CountUniqueIfColorAndValue(rangeValues As Range, criteria As Range, countRange As Range) As Long
    Dim cell As Range
    Dim i As Long
    Dim color As Long
    Dim uniqueValues As Object ' Используем Dictionary вместо Collection
    Dim value As Variant
    Dim errorMsg As String
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' Создаём Dictionary вместо Collection
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    Set ws = rangeValues.Worksheet
    color = criteria.Interior.color
    
    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).value = criteria.value And rangeValues.Cells(i).Interior.color = color Then
            value = countRange.Cells(i).value
            ' Добавляем значение в Dictionary, если его там ещё нет
            If Not uniqueValues.Exists(CStr(value)) Then
                uniqueValues.Add CStr(value), value
            End If
        End If
    Next i
    
    CountUniqueIfColorAndValue = uniqueValues.count
    Exit Function

ErrorHandler:
    errorMsg = "Произошла ошибка в функции CountUniqueIfColorAndValue:" & vbNewLine & _
               "Номер ошибки: " & Err.Number & vbNewLine & _
               "Описание: " & Err.Description & vbNewLine & _
               "Лист: " & ws.Name & vbNewLine & _
               "Ячейка: " & rangeValues.Cells(i).Address & vbNewLine & _
               "Значение в ячейке: " & value
    
    MsgBox errorMsg, vbCritical, "Ошибка в CountUniqueIfColorAndValue"
    CountUniqueIfColorAndValue = uniqueValues.count ' Возвращаем текущее количество уникальных значений
End Function

Function CountIfColorAndValue(rangeValues1 As Range, criteria1 As Range, rangeValues2 As Range, criteria2 As Range) As Long
    Dim i As Long
    Dim color1 As Long
    Dim cellValue1 As Variant, cellValue2 As Variant
    Dim countValue As Long
    Dim errorMsg As String
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = rangeValues1.Worksheet
    color1 = criteria1.Interior.color
    countValue = 0
    
    For i = 1 To rangeValues1.Cells.count
        cellValue1 = rangeValues1.Cells(i).value
        cellValue2 = rangeValues2.Cells(i).value
        If cellValue1 = criteria1.value And rangeValues1.Cells(i).Interior.color = color1 And cellValue2 = criteria2.value Then
            countValue = countValue + 1
        End If
    Next i
    
    CountIfColorAndValue = countValue
    Exit Function

ErrorHandler:
    errorMsg = "Произошла ошибка в функции CountIfColorAndValue:" & vbNewLine & _
               "Номер ошибки: " & Err.Number & vbNewLine & _
               "Описание: " & Err.Description & vbNewLine & _
               "Лист: " & ws.Name & vbNewLine & _
               "Ячейка: " & rangeValues1.Cells(i).Address
    
    MsgBox errorMsg, vbCritical, "Ошибка в CountIfColorAndValue"
    CountIfColorAndValue = countValue ' Возвращаем текущее количество
End Function

Function CountAllIfColorAndValue(rangeValues As Range, criteria As Range, countRange As Range) As Long
    Dim cell As Range
    Dim i As Long
    Dim color As Long
    Dim count As Long
    Dim errorMsg As String
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = rangeValues.Worksheet
    color = criteria.Interior.color
    count = 0
    
    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).value = criteria.value And rangeValues.Cells(i).Interior.color = color Then
            count = count + 1
        End If
    Next i
    
    CountAllIfColorAndValue = count
    Exit Function

ErrorHandler:
    errorMsg = "Произошла ошибка в функции CountAllIfColorAndValue:" & vbNewLine & _
               "Номер ошибки: " & Err.Number & vbNewLine & _
               "Описание: " & Err.Description & vbNewLine & _
               "Лист: " & ws.Name & vbNewLine & _
               "Ячейка: " & rangeValues.Cells(i).Address
    
    MsgBox errorMsg, vbCritical, "Ошибка в CountAllIfColorAndValue"
    CountAllIfColorAndValue = count ' Возвращаем текущее количество
End Function

Sub CallCombineRanges(control As IRibbonControl)
    Call CombineRanges
End Sub