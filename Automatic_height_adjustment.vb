Sub Automatic_height_adjustment()
    Dim rc As Range
    Application.ScreenUpdating = False
    For Each rc In Selection
        RowHeightForContent rc
    Next
    Application.ScreenUpdating = True
End Sub

Function RowHeightForContent(rc As Range)
    Dim MergedR_Height As Single
    Dim cellValue As String
    
    If rc.MergeCells Then
        ' Если ячейка объединена, подгоняем высоту строки под ее содержимое
        rc.WrapText = True ' Включаем перенос текста
        cellValue = rc.Value
        rc.Rows.AutoFit
        If rc.RowHeight < 25 Then
            ' Устанавливаем минимальную высоту строки в 25 пикселей (подстройте под нужные параметры)
            rc.RowHeight = 25
        End If
    ElseIf Not Application.WorksheetFunction.CountA(rc) = 0 Then
        ' Если ячейка не объединена и содержит данные, подгоняем высоту строки под ее содержимое
        rc.WrapText = True ' Включаем перенос текста
        cellValue = rc.Value
        rc.Rows.AutoFit
        If rc.RowHeight < 25 Then
            ' Устанавливаем минимальную высоту строки в 25 пикселей (подстройте под нужные параметры)
            rc.RowHeight = 25
        End If
    End If
End Function


Sub CallAutomatic_height_adjustment(control As IRibbonControl)
    Call Automatic_height_adjustment
End Sub

