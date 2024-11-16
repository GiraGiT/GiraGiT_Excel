Function Levenshtein(s1 As String, s2 As String) As Integer
    Dim i As Integer, j As Integer
    Dim d() As Integer
    Dim cost As Integer
    Dim m As Integer, n As Integer
    
    m = Len(s1)
    n = Len(s2)
    ReDim d(0 To m, 0 To n)
    
    For i = 0 To m
        d(i, 0) = i
    Next i
    
    For j = 0 To n
        d(0, j) = j
    Next j
    
    For i = 1 To m
        For j = 1 To n
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            d(i, j) = Application.WorksheetFunction.Min(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i
    
    Levenshtein = d(m, n)
End Function

Function CleanString(ByVal str As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[\s\p{P}\p{S}]"
    CleanString = regex.Replace(str, "")
    CleanString = LCase(CleanString) ' Приведение к нижнему регистру
    CleanString = Replace(CleanString, " ", "") ' Удаление пробелов внутри строки
    CleanString = Replace(CleanString, "0", "o") ' Замена похожих символов
    CleanString = Replace(CleanString, "1", "i")
    CleanString = Replace(CleanString, "5", "s")
    CleanString = Replace(CleanString, "8", "b")
    CleanString = Replace(CleanString, "3", "e")
    CleanString = Replace(CleanString, "4", "a")
    CleanString = Replace(CleanString, "6", "g")
    CleanString = Replace(CleanString, "7", "t")
    CleanString = Replace(CleanString, "9", "g")
    CleanString = Replace(CleanString, "о", "o") ' Замена кириллических символов на латинские
    CleanString = Replace(CleanString, "е", "e")
    CleanString = Replace(CleanString, "а", "a")
    CleanString = Replace(CleanString, "с", "c")
    CleanString = Replace(CleanString, "р", "p")
    CleanString = Replace(CleanString, "у", "y")
    CleanString = Replace(CleanString, "к", "k")
    CleanString = Replace(CleanString, "х", "x")
    CleanString = Replace(CleanString, "в", "b")
    CleanString = Replace(CleanString, "м", "m")
    CleanString = Replace(CleanString, "т", "t")
    CleanString = Replace(CleanString, "н", "h")
    CleanString = Replace(CleanString, "г", "g")
End Function

Function SimilarityPercentage(s1 As String, s2 As String) As Double
    Dim maxLength As Integer
    maxLength = Application.WorksheetFunction.Max(Len(s1), Len(s2))
    If maxLength = 0 Then
        SimilarityPercentage = 1
    Else
        SimilarityPercentage = (maxLength - Levenshtein(s1, s2)) / maxLength
    End If
End Function

Sub DataReconciliation()
    Dim firstRange As Range
    Dim secondRange As Range
    Dim cell As Range
    Dim matchCell As Range
    Dim minDistance As Integer
    Dim currentDistance As Integer
    Dim bestMatch As Range
    Dim differences As String
    Dim fileName1 As String
    Dim fileName2 As String
    Dim sheetName1 As String
    Dim sheetName2 As String
    Dim cellToComment As Range
    Dim processedCells As Collection
    Dim similarityThreshold As Double
    Dim mode As Variant
    
    ' Запрос режима
    mode = Application.InputBox("Выберите режим: 1 - Точные совпадения, 2 - Поиск различий", Type:=1)
    If mode = False Then Exit Sub ' Проверка на нажатие кнопки "Отмена"
    If mode <> 1 And mode <> 2 Then
        MsgBox "Неверный режим. Пожалуйста, выберите 1 или 2."
        Exit Sub
    End If
    
    If mode = 2 Then
        similarityThreshold = 0.85 ' Установите пороговое значение для процентного соотношения совпадений
    End If
    
    ' Запрос первого диапазона
    On Error Resume Next
    Set firstRange = Application.InputBox("Выделите первый диапазон ячеек:", Type:=8)
    If firstRange Is Nothing Then Exit Sub
    
    ' Запрос второго диапазона
    Set secondRange = Application.InputBox("Выделите второй диапазон ячеек:", Type:=8)
    If secondRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Получение имен файлов и листов
    fileName1 = firstRange.Worksheet.Parent.Name
    sheetName1 = firstRange.Worksheet.Name
    fileName2 = secondRange.Worksheet.Parent.Name
    sheetName2 = secondRange.Worksheet.Name
    
    ' Инициализация коллекции для отслеживания обработанных ячеек
    Set processedCells = New Collection
    
    ' Удаление существующих комментариев и очистка заливки ячеек только в первом диапазоне
    For Each cell In firstRange
        If Not cell.Comment Is Nothing Then
            cell.Comment.Delete
        End If
        cell.Interior.ColorIndex = xlNone
    Next cell
    
    ' Поиск наиболее похожих ячеек с использованием выбранного режима
    For Each cell In firstRange
        ' Проверка, является ли ячейка первой в объединённом диапазоне
        If cell.MergeCells Then
            Set cellToComment = cell.MergeArea.Cells(1, 1)
        Else
            Set cellToComment = cell
        End If
        
        ' Пропуск скрытых ячеек
        If cellToComment.EntireRow.Hidden Or cellToComment.EntireColumn.Hidden Then
            GoTo NextCell
        End If
        
        ' Пропуск ячеек, которые уже были обработаны
        On Error Resume Next
        processedCells.Add cellToComment, cellToComment.Address
        If Err.Number = 457 Then
            ' Ячейка уже была обработана
            Err.Clear
            On Error GoTo 0
            GoTo NextCell
        End If
        On Error GoTo 0
        
        ' Удаление пробелов, знаков препинания и специальных символов
        Dim cleanedCellValue As String
        cleanedCellValue = CleanString(cell.value)
        
        If mode = 1 Then
            ' Режим точных совпадений
            Dim exactMatchFound As Boolean
            exactMatchFound = False
            For Each matchCell In secondRange
                ' Пропуск скрытых ячеек
                If matchCell.EntireRow.Hidden Or matchCell.EntireColumn.Hidden Then
                    GoTo NextMatchCell
                End If
                
                If cleanedCellValue = CleanString(matchCell.value) Then
                    exactMatchFound = True
                    Exit For
                End If
                
NextMatchCell:
            Next matchCell
            
            ' Заливка ячейки в зависимости от совпадения
            If exactMatchFound Then
                cellToComment.Interior.color = RGB(0, 255, 0) ' Зелёный цвет для точного совпадения
            Else
                cellToComment.Interior.color = RGB(255, 0, 0) ' Красный цвет для ячеек без совпадений
            End If
            
        ElseIf mode = 2 Then
            ' Режим алгоритма Левенштейна
            minDistance = Application.WorksheetFunction.Max(Len(cleanedCellValue), Len(CleanString(secondRange.Cells(1, 1).value)))
            differences = ""
            Set bestMatch = Nothing
            For Each matchCell In secondRange
                ' Пропуск скрытых ячеек
                If matchCell.EntireRow.Hidden Or matchCell.EntireColumn.Hidden Then
                    GoTo NextMatchCell2
                End If
                
                currentDistance = Levenshtein(cleanedCellValue, CleanString(matchCell.value))
                If currentDistance < minDistance Then
                    minDistance = currentDistance
                    Set bestMatch = matchCell
                    differences = cell.value & vbCrLf & matchCell.value
                End If
                
NextMatchCell2:
            Next matchCell
            
            ' Проверка на пустые значения
            If cell.value = "" Or bestMatch.value = "" Then
                cellToComment.Interior.color = RGB(255, 0, 0) ' Красный цвет для ячеек без совпадений
            Else
                ' Удаление существующего комментария, если он есть
                If Not cellToComment.Comment Is Nothing Then
                    cellToComment.Comment.Delete
                End If
                
                ' Заливка ячейки в зависимости от совпадения
                If minDistance = 0 Then
                    cellToComment.Interior.color = RGB(0, 255, 0) ' Зелёный цвет для точного совпадения
                ElseIf minDistance > 0 And SimilarityPercentage(cleanedCellValue, CleanString(bestMatch.value)) >= similarityThreshold Then
                    ' Добавление комментария к ячейке
                    If Not bestMatch Is Nothing Then
                        cellToComment.AddComment Text:="Имя файла 1: " & fileName1 & " > " & sheetName1 & vbCrLf & _
                                                    "Имя файла 2: " & fileName2 & " > " & sheetName2 & vbCrLf & vbCrLf & _
                                                    "Различия:" & vbCrLf & differences
                        ' Изменение размеров окна комментария
                        With cellToComment.Comment.Shape
                            .Width = 500
                            .Height = 100
                        End With
                        
                        cellToComment.Interior.color = RGB(255, 255, 0) ' Жёлтый цвет для совпадений
                    Else
                        cellToComment.Interior.color = RGB(255, 0, 0) ' Красный цвет для ячеек без совпадений
                    End If
                ElseIf InStr(cleanedCellValue, CleanString(bestMatch.value)) > 0 Or InStr(CleanString(bestMatch.value), cleanedCellValue) > 0 Then
                    ' Проверка на частичное совпадение
                    cellToComment.Interior.color = RGB(255, 255, 0) ' Жёлтый цвет для частичных совпадений
                    ' Добавление комментария к ячейке
                    If Not bestMatch Is Nothing Then
                        cellToComment.AddComment Text:="Имя файла 1: " & fileName1 & " > " & sheetName1 & vbCrLf & _
                                                    "Имя файла 2: " & fileName2 & " > " & sheetName2 & vbCrLf & vbCrLf & _
                                                    "Различия:" & vbCrLf & differences
                        ' Изменение размеров окна комментария
                        With cellToComment.Comment.Shape
                            .Width = 500
                            .Height = 100
                        End With
                    End If
                Else
                    cellToComment.Interior.color = RGB(255, 0, 0) ' Красный цвет для ячеек без совпадений
                End If
            End If
        End If
        
        ' Отладочное сообщение
        If Not cellToComment.Comment Is Nothing Then
            Debug.Print "Ячейка: " & cellToComment.Address & " - Комментарий: " & cellToComment.Comment.Text
        Else
            Debug.Print "Ячейка: " & cellToComment.Address & " - Комментарий не добавлен"
        End If
        
NextCell:
    Next cell
End Sub

Sub CallDataReconciliation(control As IRibbonControl)
    Call DataReconciliation
End Sub