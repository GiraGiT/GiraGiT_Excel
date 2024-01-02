Sub Automatic_height_adjustment()

    Dim rng As Range
    Dim cell As Range
    Dim txtBox As Shape

    ' Используем текущий выделенный диапазон
    Set rng = Selection

    ' Проходим по каждой ячейке в диапазоне
    For Each cell In rng

        ' Проверяем, является ли ячейка частью объединенной ячейки
        If cell.MergeCells Then

            ' Создаем текстовое поле в ячейке
            Set txtBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, cell.Left, cell.Top, cell.Width, cell.Height)

            ' Копируем текст из ячейки в текстовое поле
            txtBox.TextFrame.Characters.Text = cell.Value

            ' Устанавливаем свойства текстового поля
            With txtBox.TextFrame
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .AutoSize = True
            End With

            ' Устанавливаем высоту ячейки, чтобы соответствовать высоте текстового поля
            cell.RowHeight = txtBox.Height

            ' Удаляем текстовое поле
            txtBox.Delete

        End If

    Next cell

End Sub

Sub CallAutomatic_height_adjustment(control As IRibbonControl)
    Call Automatic_height_adjustment
End Sub
