Sub RefreshSheetFormulas()
    Dim rng As Range
    Dim cell As Range

    ' Отключить некоторые функции Excel
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each cell In ActiveSheet.UsedRange
        If cell.HasFormula Then
            cell.Formula = cell.Formula
        End If
    Next cell

    ' Включить функции Excel обратно
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


Sub CallRefreshSheetFormulas(control As IRibbonControl)
    Call RefreshSheetFormulas
    MsgBox "Ôîðìóëû íà òåêóùåì ëèñòå îáíîâëåíû", vbInformation
End Sub
