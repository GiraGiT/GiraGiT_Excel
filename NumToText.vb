Option Explicit
Public Function NumToText(chislo As Double, Optional rod As Integer = 1) As String
    Dim drobLen As Variant, x2 As Double, utr As String, utr1, ITOG0 As String, itog As String
    Dim tekst As String, sklon As String, drobprop As String
    
    If chislo > 1E+15 Then
        NumToText = "Аргумент больше 999 999 999 999 999.99!"
        Exit Function
    ElseIf chislo < 0 Then
        NumToText = "Аргумент отрицательный!"
        Exit Function
    End If

    ' Проверка на наличие дробной части
    If chislo = Int(chislo) Then
        ' Целое число
        chislo = Int(chislo)
        ITOG0 = Trim(СуммаOPC(chislo, 1, rod))
    Else
        ' Дробное число
        Dim целаяЧасть As Double, дробнаяЧасть As Double
        целаяЧасть = Int(chislo)
        дробнаяЧасть = CDbl(Mid(chislo, InStr(1, chislo, ",") + 1))
        
        utr = str(chislo)
        utr1 = Split(utr, ".")
        drobLen = Len(utr1(1))
        
        itog = Trim(СуммаOPC(целаяЧасть, 3, 1)) ' Мужской род для целой части
        drobprop = Trim(СуммаOPC(Int(дробнаяЧасть), 3, 1)) ' Мужской род для дробной части
        
        ' Правильное склонение "целых/целая"
        If целаяЧасть = 1 Then
            tekst = " целая "
        Else
            tekst = " целых "
        End If
        
        sklon = drobnaya(дробнаяЧасть, drobLen)
        ITOG0 = Trim(itog & tekst & drobprop & " " & sklon)
    End If

    ' Преобразование первой буквы в заглавную и удаление лишних пробелов
    NumToText = UCase(Left(ITOG0, 1)) & Mid(ITOG0, 2)
End Function

Public Function СуммаOPC(x As Double, opc As Long, rod As Integer) As String
    Dim y(1 To 5) As Integer, i1 As Byte
    Dim Text(1 To 5) As String, i2 As Byte, y1 As Byte, y2 As Byte, _
            y3 As Byte, Text0 As String, text1 As String, text2 As String, Text3 As String, _
            Text4 As String, Text5 As String
    For i1 = 1 To 5
        x = Fix(x) / 1000
        y(i1) = (x - Fix(x)) * 1000
    Next
    For i2 = 1 To 5
        y1 = y(i2) Mod 10
        y2 = (y(i2) - y1) / 10 Mod 10
        y3 = y(i2) \ 100
        text1 = Choose(y3 + 1, "", "сто ", "двести ", "триста ", "четыреста ", _
                "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
        text2 = Choose(y2 + 1, "", "", "двадцать ", "тридцать ", "сорок ", _
                "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
        If y2 = 1 Then
            Text3 = Choose(y1 + 1, "десять ", "одиннадцать ", "двенадцать ", _
                    "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", _
                    "семнадцать ", "восемнадцать ", "девятнадцать ")
        Else
            If y2 <> 1 Then
                Text3 = Choose(y1 + 1, "", IIf(i2 = 2, "одна ", "один "), "две ", "три ", "четыре ", "пять ", _
                        "шесть ", "семь ", "восемь ", "девять ")
            Else
                Text3 = Choose(y1 + 1, "", IIf(rod = 2, "одна ", "один "), "один ", "два ", "три ", "четыре ", "пять ", _
                        "шесть ", "семь ", "восемь ", "девять ")
            End If
        End If
        If y2 <> 1 And y1 = 1 Then
            Text4 = Choose(i2, "", "тысяча ", "миллион ", "миллиард ", "триллион ")
        ElseIf y2 <> 1 And y1 > 1 And y1 < 5 Then
            Text4 = Choose(i2, "", "тысячи ", "миллиона ", "миллиарда ", "триллиона ")
        ElseIf y1 = 0 And y2 = 0 And y3 = 0 Then
            Text4 = Choose(i2, "", "", "", "", "")
        Else
            Text4 = Choose(i2, "", "тысяч ", "миллионов ", "миллиардов ", "триллионов ")
        End If
        Text(i2) = text1 & text2 & Text3 & Text4
    Next
    If y(1) + y(2) + y(3) + y(4) + y(5) = 0 Then
        Text0 = "ноль "
    Else
        Text0 = Text(5) & Text(4) & Text(3) & Text(2) & Text(1)
    End If
    СуммаOPC = Text0
End Function

Public Function drobnaya(x2 As Double, drobLen As Variant) As String
    Dim x As Variant, scl As String
    
    ' Максимальная длина - триллионные
    If drobLen > 12 Then
        drobLen = 12
    End If
    
    x = Right(x2, 1)
    If x = 1 Then
        scl = Choose(drobLen, "десятая", "сотая", "тысячная", "десятитысячная", _
                "стотысячная", "миллионная", "десятимиллионная", "стомиллионная", _
                "миллиардная", "десятимиллиардная", "стомиллиардная", "триллионная")
    Else
        scl = Choose(drobLen, "десятых", "сотых", "тысячных", "десятитысячных", _
                "стотысячных", "миллионных", "десятимиллионных", "стомиллионных", _
                "миллиардных", "десятимиллиардных", "стомиллиардных", "триллионных")
    End If
    drobnaya = scl
End Function

Sub InsertFormula()
    Dim первоеЗначение As Range
    Dim формула As String
    Dim rod As Integer
    
    ' Выбор ячейки для первого значения
    On Error Resume Next ' Подавляем ошибку, если не выбрана ячейка
    Set первоеЗначение = Application.InputBox("Выберите ячейку для первого значения:", Type:=8)
    On Error GoTo 0 ' Возвращаем стандартный обработчик ошибок
    
    If первоеЗначение Is Nothing Then Exit Sub ' Если не выбрана ячейка, выходим из макроса

    ' Проверка на наличие дробной части
    If первоеЗначение.value = Int(первоеЗначение.value) Then
        ' Выбор рода для прописи целых чисел
        rod = Application.InputBox("Выберите род для прописи целых чисел: 1 - мужской, 2 - женский", Type:=1)
        If rod < 1 Or rod > 2 Then Exit Sub ' Если род выбран неверно, выходим из макроса
    Else
        rod = 1 ' По умолчанию мужской род для дробных чисел
    End If

    ' Формирование строки формулы
    формула = "=NumToText(" & первоеЗначение.Address & "," & rod & ")"
    
    ' Изменение формата активной ячейки на "Общий"
    ActiveCell.NumberFormat = "General"
    
    ' Вставка формулы в активную ячейку
    ActiveCell.formula = формула
End Sub

Sub CallInsertFormula(control As IRibbonControl)
    Call InsertFormula
End Sub