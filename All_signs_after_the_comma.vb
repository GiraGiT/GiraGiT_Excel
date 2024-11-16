Sub All_signs_after_the_comma()
    ' Проверка, есть ли выделенные ячейки
    If Selection.count = 0 Then
        MsgBox "Нет выделенных ячеек", vbInformation
        Exit Sub
    End If

    ' Применение пользовательского формата к выделенным ячейкам
    For Each cell In Selection
        If IsNumeric(cell.value) And (InStr(cell.value, ".") > 0 Or InStr(cell.value, ",") > 0) Then
            cell.NumberFormat = "0.##########################################################"
        Else
            cell.NumberFormat = "General"
        End If
    Next cell
End Sub

Sub CallAll_signs_after_the_comma(control As IRibbonControl)
    Call All_signs_after_the_comma
End Sub