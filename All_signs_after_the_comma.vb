Sub All_signs_after_the_comma()
    ' Проверка, есть ли выделенные ячейки
    If Selection.Count = 0 Then
        MsgBox "Нет выделенных ячеек", vbInformation
        Exit Sub
    End If

    ' Применение пользовательского формата к выделенным ячейкам
    Selection.NumberFormat = "0.#############################"
End Sub


Sub CallAll_signs_after_the_comma(control As IRibbonControl)
    Call All_signs_after_the_comma
End Sub