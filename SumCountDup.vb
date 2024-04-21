Sub CombineRanges()
    Dim choice As Variant
    Dim rangeValues As Range
    Dim criteria As Range
    Dim rangeSum As Range
    Dim countRange As Range
    Dim criteria2 As Range

    choice = InputBox("Выберите действие: введите '1' для суммирования, '2' - для подсчета")

    If choice = "" Then Exit Sub

    On Error Resume Next
    Set rangeValues = Application.InputBox("Выберите диапазон для поиска значения и цвета", Type:=8)
    Set criteria = Application.InputBox("Выберите ячейку с значением и цветом для поиска", Type:=8)
    On Error GoTo 0

    If choice = "1" Then
        On Error Resume Next
        Set rangeSum = Application.InputBox("Выберите диапазон для суммирования", Type:=8)
        On Error GoTo 0
        ActiveCell.formula = "=SumIfColorAndValue(" & rangeValues.Address & ", " & criteria.Address & ", " & rangeSum.Address & ")"
    ElseIf choice = "2" Then
        On Error Resume Next
        Set countRange = Application.InputBox("Выберите диапазон для подсчета", Type:=8)
        Set criteria2 = Application.InputBox("Выберите ячейку с значением для подсчета", Type:=8)
        On Error GoTo 0
        ActiveCell.formula = "=CountIfColorAndValue(" & rangeValues.Address & ", " & criteria.Address & ", " & countRange.Address & ", " & criteria2.Address & ")"
    End If
End Sub

Function SumIfColorAndValue(rangeValues As Range, criteria As Range, rangeSum As Range) As Double
    Dim cell As Range
    Dim i As Long
    Dim color As Long

    color = criteria.Interior.color

    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).Value = criteria.Value And rangeValues.Cells(i).Interior.color = color Then
            SumIfColorAndValue = SumIfColorAndValue + rangeSum.Cells(i).Value
        End If
    Next i
End Function

Function CountIfColorAndValue(rangeValues As Range, criteria As Range, countRange As Range, criteria2 As Range) As Long
    Dim cell As Range
    Dim i As Long
    Dim color As Long
    Dim count As Long

    color = criteria.Interior.color

    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).Value = criteria.Value And rangeValues.Cells(i).Interior.color = color Then
            If countRange.Cells(i).Value = criteria2.Value Then
                count = count + 1
            End If
        End If
    Next i

    CountIfColorAndValue = count
End Function

Sub CallCombineRanges(control As IRibbonControl)
    Call CombineRanges
End Sub