Sub InsertRowsColumns()
    Dim rng As Range
    Dim count As Variant

    On Error GoTo ErrorHandler ' Добавляем обработчик ошибок

    ' Получите выделенный диапазон
    Set rng = Selection

    ' Запросите у пользователя количество строк или столбцов для вставки
    count = Application.InputBox("Введите количество строк или столбцов для вставки:", Type:=1)

    ' Если пользователь нажал "Отмена", прекратить выполнение
    If count = False Then Exit Sub

    ' Если пользователь нажал "ОК" без ввода значения или ввел значение, начинающееся с "=" или "-", показать сообщение об ошибке и прекратить выполнение
    If count = "" Or Left(count, 1) = "=" Or Left(count, 1) = "-" Then
        MsgBox "Введено недопустимое значение.", vbInformation, "Информация"
        Exit Sub
    End If

    ' Вставьте строки или столбцы
    If rng.Rows.count >= rng.Columns.count Then
        rng.Resize(rng.Rows.count, count).EntireColumn.Insert
    Else
        rng.Resize(count, rng.Columns.count).EntireRow.Insert
    End If

    ' Обновите формулы
    Call UpdateFormulas

    Exit Sub ' Выходим из подпрограммы, если нет ошибок

ErrorHandler: ' Обработка ошибок
    If Err.Number <> 13 Then ' Если ошибка не связана с отменой ввода
        MsgBox "Произошла ошибка: " & Err.Description, vbCritical, "Ошибка"
    End If
End Sub

Sub UpdateFormulas()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    ' Пройдите по всем формулам на текущем листе
    Set ws = ActiveSheet
    On Error Resume Next
    Set rng = ws.Cells.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not rng Is Nothing Then
        For Each cell In rng
            ' Обновите формулу
            cell.Formula = cell.Formula
        Next cell
    End If
End Sub

Sub CallInsertRowsColumns(control As IRibbonControl)
    On Error GoTo ErrorHandler ' Добавляем обработчик ошибок

    Call InsertRowsColumns

    Exit Sub ' Выходим из подпрограммы, если нет ошибок

ErrorHandler: ' Обработка ошибок
    If Err.Number <> 13 Then ' Если ошибка не связана с отменой ввода
        MsgBox "Произошла ошибка: " & Err.Description, vbCritical, "Ошибка"
    End If
End Sub