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
    CleanString = LCase(CleanString) ' Ïðèâåäåíèå ê íèæíåìó ðåãèñòðó
    CleanString = Replace(CleanString, " ", "") ' Óäàëåíèå ïðîáåëîâ âíóòðè ñòðîêè
    CleanString = Replace(CleanString, "0", "o") ' Çàìåíà ïîõîæèõ ñèìâîëîâ
    CleanString = Replace(CleanString, "1", "i")
    CleanString = Replace(CleanString, "5", "s")
    CleanString = Replace(CleanString, "8", "b")
    CleanString = Replace(CleanString, "3", "e")
    CleanString = Replace(CleanString, "4", "a")
    CleanString = Replace(CleanString, "6", "g")
    CleanString = Replace(CleanString, "7", "t")
    CleanString = Replace(CleanString, "9", "g")
    CleanString = Replace(CleanString, "î", "o") ' Çàìåíà êèðèëëè÷åñêèõ ñèìâîëîâ íà ëàòèíñêèå
    CleanString = Replace(CleanString, "å", "e")
    CleanString = Replace(CleanString, "à", "a")
    CleanString = Replace(CleanString, "ñ", "c")
    CleanString = Replace(CleanString, "ð", "p")
    CleanString = Replace(CleanString, "ó", "y")
    CleanString = Replace(CleanString, "ê", "k")
    CleanString = Replace(CleanString, "õ", "x")
    CleanString = Replace(CleanString, "â", "b")
    CleanString = Replace(CleanString, "ì", "m")
    CleanString = Replace(CleanString, "ò", "t")
    CleanString = Replace(CleanString, "í", "h")
    CleanString = Replace(CleanString, "ã", "g")
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
    
    ' Çàïðîñ ðåæèìà
    mode = Application.InputBox("Âûáåðèòå ðåæèì: 1 - Òî÷íûå ñîâïàäåíèÿ, 2 - Ïîèñê ðàçëè÷èé", Type:=1)
    If mode = False Then Exit Sub ' Ïðîâåðêà íà íàæàòèå êíîïêè "Îòìåíà"
    If mode <> 1 And mode <> 2 Then
        MsgBox "Íåâåðíûé ðåæèì. Ïîæàëóéñòà, âûáåðèòå 1 èëè 2."
        Exit Sub
    End If
    
    If mode = 2 Then
        similarityThreshold = 0.85 ' Óñòàíîâèòå ïîðîãîâîå çíà÷åíèå äëÿ ïðîöåíòíîãî ñîîòíîøåíèÿ ñîâïàäåíèé
    End If
    
    ' Çàïðîñ ïåðâîãî äèàïàçîíà
    On Error Resume Next
    Set firstRange = Application.InputBox("Âûäåëèòå ïåðâûé äèàïàçîí ÿ÷ååê:", Type:=8)
    If firstRange Is Nothing Then Exit Sub
    
    ' Çàïðîñ âòîðîãî äèàïàçîíà
    Set secondRange = Application.InputBox("Âûäåëèòå âòîðîé äèàïàçîí ÿ÷ååê:", Type:=8)
    If secondRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Ïîëó÷åíèå èìåí ôàéëîâ è ëèñòîâ
    fileName1 = firstRange.Worksheet.Parent.Name
    sheetName1 = firstRange.Worksheet.Name
    fileName2 = secondRange.Worksheet.Parent.Name
    sheetName2 = secondRange.Worksheet.Name
    
    ' Èíèöèàëèçàöèÿ êîëëåêöèè äëÿ îòñëåæèâàíèÿ îáðàáîòàííûõ ÿ÷ååê
    Set processedCells = New Collection
    
    ' Óäàëåíèå ñóùåñòâóþùèõ êîììåíòàðèåâ è î÷èñòêà çàëèâêè ÿ÷ååê òîëüêî â ïåðâîì äèàïàçîíå
    For Each cell In firstRange
        If Not cell.Comment Is Nothing Then
            cell.Comment.Delete
        End If
        cell.Interior.ColorIndex = xlNone
    Next cell
    
    ' Ïîèñê íàèáîëåå ïîõîæèõ ÿ÷ååê ñ èñïîëüçîâàíèåì âûáðàííîãî ðåæèìà
    For Each cell In firstRange
        ' Ïðîâåðêà, ÿâëÿåòñÿ ëè ÿ÷åéêà ïåðâîé â îáúåäèí¸ííîì äèàïàçîíå
        If cell.MergeCells Then
            Set cellToComment = cell.MergeArea.Cells(1, 1)
        Else
            Set cellToComment = cell
        End If
        
        ' Ïðîïóñê ñêðûòûõ ÿ÷ååê
        If cellToComment.EntireRow.Hidden Or cellToComment.EntireColumn.Hidden Then
            GoTo NextCell
        End If
        
        ' Ïðîïóñê ÿ÷ååê, êîòîðûå óæå áûëè îáðàáîòàíû
        On Error Resume Next
        processedCells.Add cellToComment, cellToComment.Address
        If Err.Number = 457 Then
            ' ß÷åéêà óæå áûëà îáðàáîòàíà
            Err.Clear
            On Error GoTo 0
            GoTo NextCell
        End If
        On Error GoTo 0
        
        ' Óäàëåíèå ïðîáåëîâ, çíàêîâ ïðåïèíàíèÿ è ñïåöèàëüíûõ ñèìâîëîâ
        Dim cleanedCellValue As String
        cleanedCellValue = CleanString(cell.value)
        
        If mode = 1 Then
            ' Ðåæèì òî÷íûõ ñîâïàäåíèé
            Dim exactMatchFound As Boolean
            exactMatchFound = False
            For Each matchCell In secondRange
                ' Ïðîïóñê ñêðûòûõ ÿ÷ååê
                If matchCell.EntireRow.Hidden Or matchCell.EntireColumn.Hidden Then
                    GoTo NextMatchCell
                End If
                
                If cleanedCellValue = CleanString(matchCell.value) Then
                    exactMatchFound = True
                    Exit For
                End If
                
NextMatchCell:
            Next matchCell
            
            ' Çàëèâêà ÿ÷åéêè â çàâèñèìîñòè îò ñîâïàäåíèÿ
            If exactMatchFound Then
                cellToComment.Interior.color = RGB(0, 255, 0) ' Çåë¸íûé öâåò äëÿ òî÷íîãî ñîâïàäåíèÿ
            Else
                cellToComment.Interior.color = RGB(255, 0, 0) ' Êðàñíûé öâåò äëÿ ÿ÷ååê áåç ñîâïàäåíèé
            End If
            
        ElseIf mode = 2 Then
            ' Ðåæèì àëãîðèòìà Ëåâåíøòåéíà
            minDistance = Application.WorksheetFunction.Max(Len(cleanedCellValue), Len(CleanString(secondRange.Cells(1, 1).value)))
            differences = ""
            Set bestMatch = Nothing
            For Each matchCell In secondRange
                ' Ïðîïóñê ñêðûòûõ ÿ÷ååê
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
            
            ' Ïðîâåðêà íà ïóñòûå çíà÷åíèÿ
            If cell.value = "" Or bestMatch.value = "" Then
                cellToComment.Interior.color = RGB(255, 0, 0) ' Êðàñíûé öâåò äëÿ ÿ÷ååê áåç ñîâïàäåíèé
            Else
                ' Óäàëåíèå ñóùåñòâóþùåãî êîììåíòàðèÿ, åñëè îí åñòü
                If Not cellToComment.Comment Is Nothing Then
                    cellToComment.Comment.Delete
                End If
                
                ' Çàëèâêà ÿ÷åéêè â çàâèñèìîñòè îò ñîâïàäåíèÿ
                If minDistance = 0 Then
                    cellToComment.Interior.color = RGB(0, 255, 0) ' Çåë¸íûé öâåò äëÿ òî÷íîãî ñîâïàäåíèÿ
                ElseIf minDistance > 0 And SimilarityPercentage(cleanedCellValue, CleanString(bestMatch.value)) >= similarityThreshold Then
                    ' Äîáàâëåíèå êîììåíòàðèÿ ê ÿ÷åéêå
                    If Not bestMatch Is Nothing Then
                        cellToComment.AddComment Text:="Èìÿ ôàéëà 1: " & fileName1 & " > " & sheetName1 & vbCrLf & _
                                                    "Èìÿ ôàéëà 2: " & fileName2 & " > " & sheetName2 & vbCrLf & vbCrLf & _
                                                    "Ðàçëè÷èÿ:" & vbCrLf & differences
                        ' Èçìåíåíèå ðàçìåðîâ îêíà êîììåíòàðèÿ
                        With cellToComment.Comment.Shape
                            .Width = 500
                            .Height = 100
                        End With
                        
                        cellToComment.Interior.color = RGB(255, 255, 0) ' Æ¸ëòûé öâåò äëÿ ñîâïàäåíèé
                    Else
                        cellToComment.Interior.color = RGB(255, 0, 0) ' Êðàñíûé öâåò äëÿ ÿ÷ååê áåç ñîâïàäåíèé
                    End If
                ElseIf InStr(cleanedCellValue, CleanString(bestMatch.value)) > 0 Or InStr(CleanString(bestMatch.value), cleanedCellValue) > 0 Then
                    ' Ïðîâåðêà íà ÷àñòè÷íîå ñîâïàäåíèå
                    cellToComment.Interior.color = RGB(255, 255, 0) ' Æ¸ëòûé öâåò äëÿ ÷àñòè÷íûõ ñîâïàäåíèé
                    ' Äîáàâëåíèå êîììåíòàðèÿ ê ÿ÷åéêå
                    If Not bestMatch Is Nothing Then
                        cellToComment.AddComment Text:="Èìÿ ôàéëà 1: " & fileName1 & " > " & sheetName1 & vbCrLf & _
                                                    "Èìÿ ôàéëà 2: " & fileName2 & " > " & sheetName2 & vbCrLf & vbCrLf & _
                                                    "Ðàçëè÷èÿ:" & vbCrLf & differences
                        ' Èçìåíåíèå ðàçìåðîâ îêíà êîììåíòàðèÿ
                        With cellToComment.Comment.Shape
                            .Width = 500
                            .Height = 100
                        End With
                    End If
                Else
                    cellToComment.Interior.color = RGB(255, 0, 0) ' Êðàñíûé öâåò äëÿ ÿ÷ååê áåç ñîâïàäåíèé
                End If
            End If
        End If
        
        ' Îòëàäî÷íîå ñîîáùåíèå
        If Not cellToComment.Comment Is Nothing Then
            Debug.Print "ß÷åéêà: " & cellToComment.Address & " - Êîììåíòàðèé: " & cellToComment.Comment.Text
        Else
            Debug.Print "ß÷åéêà: " & cellToComment.Address & " - Êîììåíòàðèé íå äîáàâëåí"
        End If
        
NextCell:
    Next cell
End Sub

Sub CallDataReconciliation(control As IRibbonControl)
    Call DataReconciliation
End Sub
