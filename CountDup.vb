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