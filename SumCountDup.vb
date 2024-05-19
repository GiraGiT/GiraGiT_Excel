Sub CombineRanges()
    Dim choice As Variant
    Dim rangeValues As Range
    Dim criteria As Range
    Dim rangeSum As Range
    Dim rangeValues2 As Range
    Dim criteria2 As Range
    Dim countRange As Range

    choice = InputBox("Выберите действие: введите '1' для суммирования, '2' - для суммирования по двум критериям, '3' - для подсчета уникальных значений, '4' - для подсчета значений по цвету, '5' - для подсчета всех значений по цвету")

    If choice = "" Then Exit Sub

    On Error Resume Next
    Set rangeValues = Application.InputBox("Выберите диапазон для поиска значения и цвета", Type:=8)
    Set criteria = Application.InputBox("Выберите ячейку с значением и цветом для поиска", Type:=8)
    On Error GoTo 0

    If choice = "1" Then
        On Error Resume Next
        Set rangeSum = Application.InputBox("Выберите диапазон для суммирования", Type:=8)
        On Error GoTo 0
        ActiveCell.Formula = "=SumIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & rangeSum.Columns(1).Address & ")"
    ElseIf choice = "2" Then
        On Error Resume Next
        Set rangeValues2 = Application.InputBox("Выберите второй диапазон для поиска значения и цвета", Type:=8)
        Set criteria2 = Application.InputBox("Выберите вторую ячейку с значением и цветом для поиска", Type:=8)
        Set rangeSum = Application.InputBox("Выберите диапазон для суммирования", Type:=8)
        On Error GoTo 0
        If rangeValues2 Is Nothing Or criteria2 Is Nothing Or rangeSum Is Nothing Then Exit Sub
        ActiveCell.Formula = "=SumIfColorAndValue2(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & rangeSum.Columns(1).Address & ", " & rangeValues2.Columns(1).Address & ", " & criteria2.Address & ")"
    ElseIf choice = "3" Then
        On Error Resume Next
        Set countRange = Application.InputBox("Выберите диапазон для подсчета уникальных значений", Type:=8)
        On Error GoTo 0
        ActiveCell.Formula = "=CountUniqueIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & countRange.Columns(1).Address & ")"
    ElseIf choice = "4" Then
        On Error Resume Next
        Set rangeValues2 = Application.InputBox("Выберите второй диапазон для поиска значения и цвета", Type:=8)
        Set criteria2 = Application.InputBox("Выберите вторую ячейку с значением и цветом для поиска", Type:=8)
        On Error GoTo 0
        If rangeValues2 Is Nothing Or criteria2 Is Nothing Then Exit Sub
        ActiveCell.Formula = "=CountIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & rangeValues2.Columns(1).Address & ", " & criteria2.Address & ")"
    ElseIf choice = "5" Then
        On Error Resume Next
        Set countRange = Application.InputBox("Выберите диапазон для подсчета всех значений", Type:=8)
        On Error GoTo 0
        ActiveCell.Formula = "=CountAllIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & countRange.Columns(1).Address & ")"
    End If
End Sub


Function SumIfColorAndValue(rangeValues As Range, criteria As Range, rangeSum As Range) As Double
    Dim cell As Range
    Dim i As Long
    Dim color As Long

    color = criteria.Interior.color

    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).value = criteria.value And rangeValues.Cells(i).Interior.color = color Then
            If IsNumeric(rangeSum.Cells(i).value) Then
                SumIfColorAndValue = SumIfColorAndValue + rangeSum.Cells(i).value
            End If
        End If
    Next i
End Function


Function SumIfColorAndValue2(rangeValues1 As Range, criteria1 As Range, rangeSum As Range, rangeValues2 As Range, criteria2 As Range) As Double
    Dim i As Long
    Dim color1 As Long
    Dim cellValue1 As Variant, cellValue2 As Variant
    Dim sumValue As Variant

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
End Function



Function CountUniqueIfColorAndValue(rangeValues As Range, criteria As Range, countRange As Range) As Long
    Dim cell As Range
    Dim i As Long
    Dim color As Long
    Dim uniqueValues As Collection
    Dim value As Variant

    Set uniqueValues = New Collection

    color = criteria.Interior.color

    On Error Resume Next
    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).value = criteria.value And rangeValues.Cells(i).Interior.color = color Then
            value = countRange.Cells(i).value
            uniqueValues.Add value, CStr(value)
        End If
    Next i
    On Error GoTo 0

    CountUniqueIfColorAndValue = uniqueValues.count
End Function

Function CountIfColorAndValue(rangeValues1 As Range, criteria1 As Range, rangeValues2 As Range, criteria2 As Range) As Long
    Dim i As Long
    Dim color1 As Long
    Dim cellValue1 As Variant, cellValue2 As Variant
    Dim countValue As Long

    color1 = criteria1.Interior.color

    For i = 1 To rangeValues1.Cells.count
        cellValue1 = rangeValues1.Cells(i).value
        cellValue2 = rangeValues2.Cells(i).value
        If cellValue1 = criteria1.value And rangeValues1.Cells(i).Interior.color = color1 And cellValue2 = criteria2.value Then
            countValue = countValue + 1
        End If
    Next i

    CountIfColorAndValue = countValue
End Function

Function CountAllIfColorAndValue(rangeValues As Range, criteria As Range, countRange As Range) As Long
    Dim cell As Range
    Dim i As Long
    Dim color As Long
    Dim count As Long

    color = criteria.Interior.color
    count = 0

    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).value = criteria.value And rangeValues.Cells(i).Interior.color = color Then
            count = count + 1
        End If
    Next i

    CountAllIfColorAndValue = count
End Function




Sub CallCombineRanges(control As IRibbonControl)
    Call CombineRanges
End Sub