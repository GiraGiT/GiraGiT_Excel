Sub InsertRowsColumns()
    Dim rng As Range
    Dim count As Integer

    ' Получите выделенный диапазон
    Set rng = Selection

    ' Запросите у пользователя количество строк или столбцов для вставки
    count = InputBox("Введите количество строк или столбцов для вставки:")

    ' Вставьте строки или столбцы
    If rng.Rows.count >= rng.Columns.count Then
        rng.Resize(rng.Rows.count, count).EntireColumn.Insert
    Else
        rng.Resize(count, rng.Columns.count).EntireRow.Insert
    End If
End Sub




Sub CallInsertRowsColumns(control As IRibbonControl)
    Call InsertRowsColumns
End Sub
