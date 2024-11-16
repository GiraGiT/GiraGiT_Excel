Sub AdjustRowHeight()
    Dim selectedRange As Range
    Dim rowRange As Range
    Dim textBox As Shape
    Dim cellText As String

    ' Отключаем обновление экрана и автоматический расчет формул
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Получаем выделенный диапазон
    Set selectedRange = Selection

    ' Обрабатываем каждую строку в выделенном диапазоне
    For Each rowRange In selectedRange.Rows
        ' Создаем текстовое поле и выравниваем его по верхнему краю строки и ширине столбцов в выделенном диапазоне
        Set textBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, rowRange.Left, rowRange.Top, rowRange.Width, 100)

        ' Устанавливаем свойство AutoSize в msoAutoSizeShapeToFitText для автоматической подгонки размера текстового поля под текст
        textBox.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        
        textBox.TextFrame2.MarginTop = 5
        textBox.TextFrame2.MarginLeft = 3
        textBox.TextFrame2.MarginRight = 3
        textBox.TextFrame2.MarginBottom = 5

        ' Копируем значения из первого столбца выделенного диапазона в текстовое поле
        cellText = rowRange.Cells(1, 1).value
        textBox.TextFrame2.TextRange.Characters.Text = cellText

        cellText = rowRange.Cells(1, 1).Font.Name
        textBox.TextFrame2.TextRange.Characters.Font.Name = cellText
 
        cellText = rowRange.Cells(1, 1).Font.Size
        textBox.TextFrame2.TextRange.Characters.Font.Size = cellText

        ' Выравниваем нижнюю границу выделенного диапазона по нижней границе текстового поля
        rowRange.RowHeight = textBox.Top + textBox.Height - rowRange.Top

        ' Удаляем текстовое поле после использования
        textBox.Delete
    Next rowRange

    ' Включаем обратно обновление экрана и автоматический расчет формул
    Application.ScreenUpdating = True
End Sub

Sub Automatic_height_adjustment()
    Dim ws As Worksheet
    Dim rc As Range
    Dim maxTextHeight As Single

    Set ws = ActiveSheet
    maxTextHeight = 0

    ' Проверяем, что есть выделенные ячейки
    If Not Selection Is Nothing Then
        For Each rc In Selection
            If Not IsEmpty(rc) Then
                Call AdjustRowHeight ' Вызываем функцию для подгонки высоты
                ' Обновляем максимальную высоту текста
                If rc.RowHeight > maxTextHeight Then
                    maxTextHeight = rc.RowHeight
                End If
                ' Центрируем содержимое по вертикали
                rc.VerticalAlignment = xlCenter
            End If
        Next rc
    End If
End Sub

Sub CallAutomatic_height_adjustment(control As IRibbonControl)
    Call Automatic_height_adjustment
End Sub
