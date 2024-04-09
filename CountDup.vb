Sub summ_or_dup_by_color()
    Dim rng As Range
    Dim arg As Range
    Dim funcNum As Integer
    Dim formula As String
    
    ' Пользователь выбирает диапазон
    Set rng = Application.InputBox("Выберите диапазон", Type:=8)
    
    ' Пользователь выбирает функцию
    funcNum = Application.InputBox("Выберите функцию: 1 - Сумма по цвету, 2 - Подсчет по цвету", Type:=1)
    
    ' Пользователь выбирает аргумент
    Set arg = Application.InputBox("Выберите ячейку-аргумент", Type:=8)
    
    ' Создание формулы в зависимости от выбранной функции
    If funcNum = 1 Then
        formula = "=SumByColor(" & rng.Address & ", " & arg.Address & ")"
    ElseIf funcNum = 2 Then
        formula = "=countdup(" & rng.Address & ", " & arg.Address & ")"
    Else
        MsgBox "Неверный номер функции"
        Exit Sub
    End If
    
    ' Вставка формулы в активную ячейку
    ActiveCell.formula = formula
End Sub



Function SumByColor(Cell_Range As Range, Color_Cell As Range) As Double
    Dim cell As Range
    Dim Color_By_Numbers As Double
    Dim Target_Color As Long
    
    ' Получаем цвет из выбранной пользователем ячейки
    Target_Color = Color_Cell.Interior.Color
    
    ' Проходим по всем ячейкам в диапазоне
    For Each cell In Cell_Range
        If (cell.Interior.Color = Target_Color) Then
            Color_By_Numbers = Color_By_Numbers + cell.Value
        End If
    Next cell
    
    SumByColor = Color_By_Numbers
End Function

Function countdup(rng As Range, arg As Variant) As Long
    Dim cell As Range
    Dim colorToMatch As Long
    Dim count As Long
    Dim cellColor As Long
    
    ' Get the color of the argument cell
    colorToMatch = arg.Interior.Color
    
    ' Initialize the count
    count = 0
    
    ' Loop through each cell in the range
    For Each cell In rng
        ' Check if the cell value matches the argument
        If cell.Value = arg.Value Then
            ' Check if the cell color matches the argument color
            cellColor = cell.Interior.Color
            If cellColor = colorToMatch Then
                ' Increment the count
                count = count + 1
            End If
        End If
    Next cell
    
    ' Return the count
    countdup = count
End Function
Sub Callsumm_or_dup_by_color(control As IRibbonControl)
    Call summ_or_dup_by_color
End Sub
