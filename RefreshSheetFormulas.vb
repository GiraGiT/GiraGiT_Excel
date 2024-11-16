Sub RefreshSheetFormulas()
    Dim cell As Range
    Dim formulaText As String
    Dim pos As Integer

    ' Отключаем автоматический пересчет, обновление экрана и обработку событий
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Проходим по всем используемым ячейкам на активном листе
    For Each cell In ActiveSheet.UsedRange
        If cell.HasFormula Then
            ' Удаляем путь к надстройке из формулы
            formulaText = cell.Formula
            pos = InStr(1, formulaText, ".xlam'!")
            If pos > 0 Then
                ' Удаляем всё до и включая ".xlam'!"
                formulaText = Mid(formulaText, pos + 7)
                cell.Formula = "=" & formulaText
            End If
            ' Обновляем формулу
            cell.Formula = cell.Formula
        End If
    Next cell

    ' Включаем автоматический пересчет, обновление экрана и обработку событий обратно
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub CallRefreshSheetFormulas(control As IRibbonControl)
    Call RefreshSheetFormulas
    MsgBox "Формулы на текущем листе обновлены", vbInformation
End Sub