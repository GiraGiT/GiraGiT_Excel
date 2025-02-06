' Функция очистки и нормализации строки
Function CleanString(ByVal str As String) As String
    ' Приводим строку к верхнему регистру
    str = UCase(str)
    
    ' Заменяем латинскую 'c' на кириллическую 'с' до очистки
    str = Replace(str, "C", "С")
    str = Replace(str, "c", "С")
    
    ' Нормализуем написание слова "маслянный/масляный"
    str = Replace(str, "МАСЛЯНН", "МАСЛЯН")
    
    ' Удаляем все пробелы и знаки препинания
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    ' Удаляем все не буквенные и не цифровые символы
    regex.Pattern = "[^A-ZА-Я0-9]"
    str = regex.Replace(str, "")
    
    ' Замена похожих цифр на буквы
    Dim replacements As Variant
    replacements = Array( _
        Array("0", "O"), Array("1", "I"), Array("3", "E"), Array("4", "A"), _
        Array("5", "S"), Array("6", "G"), Array("7", "T"), Array("8", "B"), _
        Array("9", "G"), Array("2", "Z"))
    
    Dim i As Integer
    For i = 0 To UBound(replacements)
        str = Replace(str, replacements(i)(0), replacements(i)(1))
    Next i
    
    ' Замена кириллических букв на латинские аналоги
    Dim cyrToLat As Variant
    cyrToLat = Array( _
        Array("А", "A"), Array("В", "B"), Array("Е", "E"), Array("К", "K"), Array("М", "M"), _
        Array("Н", "H"), Array("О", "O"), Array("Р", "P"), Array("С", "C"), Array("Т", "T"), _
        Array("У", "Y"), Array("Х", "X"), Array("Ь", ""), Array("Ы", "I"), Array("Ё", "E"), _
        Array("И", "I"), Array("Й", "J"), Array("Д", "D"), Array("Л", "L"), Array("Ф", "F"), _
        Array("З", "Z"), Array("Ц", "C"), Array("Ч", "CH"), Array("Ш", "SH"), Array("Щ", "SCH"), _
        Array("Г", "G"), Array("П", "P"), Array("Ж", "ZH"), Array("Ю", "YU"), Array("Я", "YA"), _
        Array("Б", "B"), Array("Ь", ""), Array("Ъ", ""))
    For i = 0 To UBound(cyrToLat)
        str = Replace(str, cyrToLat(i)(0), cyrToLat(i)(1))
    Next i
    
    ' Замена латинских букв на кириллические аналоги
    Dim latToCyr As Variant
    latToCyr = Array( _
        Array("A", "А"), Array("B", "В"), Array("E", "Е"), Array("K", "К"), Array("M", "М"), _
        Array("H", "Н"), Array("O", "О"), Array("P", "Р"), Array("C", "С"), Array("T", "Т"), _
        Array("Y", "У"), Array("X", "Х"), Array("I", "И"), Array("J", "Й"), Array("G", "Г"), _
        Array("L", "Л"), Array("D", "Д"), Array("F", "Ф"), Array("Z", "З"), Array("N", "Н"), _
        Array("Q", "К"), Array("S", "С"), Array("V", "В"), Array("U", "Ю"), Array("W", "Ш"))
    For i = 0 To UBound(latToCyr)
        str = Replace(str, latToCyr(i)(0), latToCyr(i)(1))
    Next i
    
    CleanString = str
End Function

' Функция поиска наибольшей общей подстроки
Function LongestCommonSubstring(s1 As String, s2 As String) As String
    Dim lengths() As Long
    Dim maxLen As Long, endIndex As Long
    Dim i As Long, j As Long
    
    ReDim lengths(Len(s1), Len(s2))
    
    For i = 1 To Len(s1)
        For j = 1 To Len(s2)
            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                If i = 1 Or j = 1 Then
                    lengths(i, j) = 1
                Else
                    lengths(i, j) = lengths(i - 1, j - 1) + 1
                End If
                If lengths(i, j) > maxLen Then
                    maxLen = lengths(i, j)
                    endIndex = i
                End If
            Else
                lengths(i, j) = 0
            End If
        Next j
    Next i
    
    LongestCommonSubstring = Mid$(s1, endIndex - maxLen + 1, maxLen)
End Function

' Функция вычисления процента совпадения
Function SubstringSimilarity(str1 As String, str2 As String) As Double
    Dim s1 As String, s2 As String
    Dim commonSub As String
    Dim perc1 As Double, perc2 As Double
    
    s1 = CleanString(str1)
    s2 = CleanString(str2)

    ' Выводим отладочную информацию
    Debug.Print "Оригинал1: " & str1 & " | Очищенная1: " & s1
    Debug.Print "Оригинал2: " & str2 & " | Очищенная2: " & s2
    
    If Len(s1) = 0 Or Len(s2) = 0 Then
        SubstringSimilarity = 0
        Exit Function
    End If
    
    commonSub = LongestCommonSubstring(s1, s2)
    perc1 = Len(commonSub) / Len(s1)
    perc2 = Len(commonSub) / Len(s2)
    SubstringSimilarity = (perc1 + perc2) / 2
    
    ' Выводим процент схожести
    Debug.Print "Схожесть: " & SubstringSimilarity
End Function

' Основная процедура сравнения и окрашивания ячеек с добавлением комментариев
Sub DataReconciliation()
    Dim firstRange As Range, secondRange As Range
    Dim cell As Range, matchCell As Range
    Dim bestSim As Double, sim As Double
    Dim bestMatch As String, bestComment As String
    Dim threshold As Double
    Dim commentText As String
    Dim fileName1 As String, sheetName1 As String, fileName2 As String, sheetName2 As String
    
    threshold = 0.5  ' Понизили порог схожести с 0.7 до 0.5
    
    On Error Resume Next
    Set firstRange = Application.InputBox("Выделите первый диапазон ячеек:", Type:=8)
    If firstRange Is Nothing Then Exit Sub
    Set secondRange = Application.InputBox("Выделите второй диапазон ячеек:", Type:=8)
    If secondRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Получаем информацию о файлах и листах
    fileName1 = firstRange.Worksheet.Parent.Name
    sheetName1 = firstRange.Worksheet.Name
    fileName2 = secondRange.Worksheet.Parent.Name
    sheetName2 = secondRange.Worksheet.Name
    
    ' Очищаем предыдущее форматирование и комментарии
    firstRange.Interior.ColorIndex = xlNone
    firstRange.ClearComments
    
    For Each cell In firstRange
        If cell.MergeCells Then
            Set cell = cell.MergeArea.Cells(1, 1)
        End If
        If cell.EntireRow.Hidden Or cell.EntireColumn.Hidden Then GoTo NextCell
        
        bestSim = 0
        bestMatch = ""
        bestComment = ""
        
        For Each matchCell In secondRange
            If matchCell.EntireRow.Hidden Or matchCell.EntireColumn.Hidden Then GoTo NextMatch
            sim = SubstringSimilarity(cell.Text, matchCell.Text)
            If sim > bestSim Then
                bestSim = sim
                bestMatch = matchCell.Text
                bestComment = "Имя файла первого диапазона: " & fileName1 & " > Лист: " & sheetName1 & vbCrLf & _
                              "Имя файла второго диапазона: " & fileName2 & " > Лист: " & sheetName2 & vbCrLf & _
                              "Различия:" & vbCrLf & _
                              "Первое значение: " & cell.Text & vbCrLf & _
                              "Второе значение: " & matchCell.Text
End If
NextMatch:
        Next matchCell
        
        ' Окрашиваем и добавляем комментарий при необходимости
        If bestSim = 1 Then
            cell.Interior.color = RGB(0, 255, 0)  ' Зелёный
        ElseIf bestSim >= threshold Then
            cell.Interior.color = RGB(255, 255, 0)  ' Жёлтый
            commentText = fileName1 & " > " & sheetName1 & vbCrLf & _
                          fileName2 & " > " & sheetName2 & vbCrLf & vbCrLf & _
                          "Различия:" & vbCrLf & _
                          cell.Text & vbCrLf & _
                          bestMatch
            
            With cell
                If .Comment Is Nothing Then
                    .AddComment
                Else
                    .Comment.Delete
                    .AddComment
                End If
                .Comment.Text Text:=commentText
            End With
            
            cell.Comment.Shape.TextFrame.AutoSize = True
        Else
            cell.Interior.color = RGB(255, 0, 0)  ' Красный
        End If
NextCell:
    Next cell
    
    ' Вызываем функцию выделения различий в комментариях
    HighlightDifferences
End Sub

' Функция выделения различий
Sub HighlightDifferences()
    Dim ws As Worksheet
    Dim cell As Range
    Dim commentText As String
    Dim i As Integer
    
    ' Обрабатываем все комментарии на листе
    For Each ws In ThisWorkbook.Worksheets
        For Each cell In ws.UsedRange
            If Not cell.Comment Is Nothing Then
                commentText = cell.Comment.Text
                
                ' Обрабатываем каждый символ комментария
                For i = 1 To Len(commentText)
                    ' Окрашиваем символ "G" или "C" в красный цвет
                    If Mid(commentText, i, 1) = "G" Or Mid(commentText, i, 1) = "C" Then
                        With cell.Comment.Shape.TextFrame.Characters(i, 1).Font
                            .color = RGB(255, 0, 0)
                        End With
                    End If
                Next i
            End If
        Next cell
    Next ws
End Sub

Sub CallDataReconciliation(control As IRibbonControl)
    Call DataReconciliation
End Sub