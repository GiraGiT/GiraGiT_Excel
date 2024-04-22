Sub CombineRanges()
    Dim choice As Variant
    Dim rangeValues As Range
    Dim criteria As Range
    Dim rangeSum As Range
    Dim countRange As Range
    Dim criteria2 As Range
    Dim rangeValues2 As Range
    Dim criteria3 As Range

    choice = InputBox("Выберите действие: введите '1' для суммирования, '2' - для подсчета, '3' - для подсчета с двумя диапазонами")

    If choice = "" Then Exit Sub

    On Error Resume Next
    Set rangeValues = Application.InputBox("Выберите диапазон для поиска значения и цвета", Type:=8)
    Set criteria = Application.InputBox("Выберите ячейку с значением и цветом для поиска", Type:=8)
    On Error GoTo 0

    If choice = "1" Then
        On Error Resume Next
        Set rangeSum = Application.InputBox("Выберите диапазон для суммирования", Type:=8)
        On Error GoTo 0
        ActiveCell.formula = "=SumIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & rangeSum.Columns(1).Address & ")"
    ElseIf choice = "2" Then
        On Error Resume Next
        Set countRange = Application.InputBox("Выберите диапазон для подсчета", Type:=8)
        Set criteria2 = Application.InputBox("Выберите ячейку с значением для подсчета", Type:=8)
        On Error GoTo 0
        ActiveCell.formula = "=CountIfColorAndValue(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & countRange.Columns(1).Address & ", " & criteria2.Address & ")"
    ElseIf choice = "3" Then
        On Error Resume Next
        Set countRange = Application.InputBox("Выберите первый диапазон для подсчета", Type:=8)
        Set criteria2 = Application.InputBox("Выберите первую ячейку с значением для подсчета", Type:=8)
        Set rangeValues2 = Application.InputBox("Выберите второй диапазон для поиска значения и цвета", Type:=8)
        Set criteria3 = Application.InputBox("Выберите вторую ячейку с значением и цветом для поиска", Type:=8)
        On Error GoTo 0
        ActiveCell.formula = "=CountIfColorAndValue2(" & rangeValues.Columns(1).Address & ", " & criteria.Address & ", " & countRange.Columns(1).Address & ", " & criteria2.Address & ", " & rangeValues2.Columns(1).Address & ", " & criteria3.Address & ")"
    End If
End Sub


Function SumIfColorAndValue(rangeValues As Range, criteria As Range, rangeSum As Range) As Double
    Dim cell As Range
    Dim i As Long
    Dim color As Long

    color = criteria.Interior.color

    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).Value = criteria.Value And rangeValues.Cells(i).Interior.color = color Then
            If IsNumeric(rangeSum.Cells(i).Value) Then
                SumIfColorAndValue = SumIfColorAndValue + rangeSum.Cells(i).Value
            End If
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

Function CountIfColorAndValue2(rangeValues As Range, criteria As Range, countRange As Range, criteria2 As Range, rangeValues2 As Range, criteria3 As Range) As Long
    Dim cell As Range
    Dim i As Long
    Dim color As Long
    Dim color2 As Long
    Dim count As Long

    color = criteria.Interior.color
    color2 = criteria3.Interior.color

    For i = 1 To rangeValues.Cells.count
        If rangeValues.Cells(i).Value = criteria.Value And rangeValues.Cells(i).Interior.color = color Then
            If countRange.Cells(i).Value = criteria2.Value Then
                If rangeValues2.Cells(i).Value = criteria3.Value And rangeValues2.Cells(i).Interior.color = color2 Then
                    count = count + 1
                End If
            End If
        End If
    Next i

    CountIfColorAndValue2 = count
End Function


Sub CallCombineRanges(control As IRibbonControl)
    Call CombineRanges
End Sub