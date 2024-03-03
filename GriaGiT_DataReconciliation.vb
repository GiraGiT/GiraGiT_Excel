' Levenshtein Distance Function
Function Levenshtein(s1 As String, s2 As String) As Integer
    Dim i As Integer, j As Integer
    Dim l1 As Integer, l2 As Integer
    Dim dist() As Integer
    Dim minDist As Integer

    l1 = Len(s1)
    l2 = Len(s2)
    ReDim dist(l1, l2)

    ' Initialize the first row and column
    For i = 0 To l1
        dist(i, 0) = i
    Next
    For j = 0 To l2
        dist(0, j) = j
    Next

    ' Calculate the distance matrix
    For i = 1 To l1
        For j = 1 To l2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                dist(i, j) = dist(i - 1, j - 1)
            Else
                minDist = Application.WorksheetFunction.Min(dist(i - 1, j) + 1, dist(i, j - 1) + 1, dist(i - 1, j - 1) + 1)
                dist(i, j) = minDist
            End If
        Next
    Next

    Levenshtein = dist(l1, l2)
End Function

' Data Reconciliation Subroutine
Sub DataReconciliation()
    Dim rng1 As Range, rng2 As Range, cell1 As Range, cell2 As Range
    Dim threshold As Integer
    Dim fileName1 As String, fileName2 As String

    ' Allow user to select ranges
    On Error Resume Next
    Set rng1 = Application.InputBox("Выделите диапазон в первом файле", Type:=8)
    Set rng2 = Application.InputBox("Выделите диапазон во втором файле", Type:=8)
    On Error GoTo 0

    ' Check if ranges are set
    If rng1 Is Nothing Or rng2 Is Nothing Then
        MsgBox "Вы не выделили один и более диапазонов. Попробуйте ещё раз."
        Exit Sub
    End If

    ' Get file names
    fileName1 = rng1.Parent.Parent.Name
    fileName2 = rng2.Parent.Parent.Name

    ' Clear fill and comments
    rng1.Interior.Color = xlNone
    rng1.ClearComments

    ' Ask user for the threshold value
    threshold = Application.InputBox("Введите допустимое количество различий", Type:=1)

    ' First pass: mark exact matches with green color
    For Each cell1 In rng1
        If cell1.EntireRow.Hidden = False And cell1.EntireColumn.Hidden = False Then
            For Each cell2 In rng2
                If cell2.EntireRow.Hidden = False And cell2.EntireColumn.Hidden = False Then
                    If Trim(cell1.Value) <> "" And Trim(cell2.Value) <> "" And Replace(cell1.Value, " ", "") = Replace(cell2.Value, " ", "") Then
                        cell1.Interior.Color = RGB(0, 255, 0) ' Green
                        Exit For
                    End If
                End If
            Next cell2
        End If
    Next cell1

    ' Second pass: compare non-matching cells based on Levenshtein distance
    For Each cell1 In rng1
        ' Ignore matched cells and empty cells
        If cell1.Interior.Color <> RGB(0, 255, 0) And Trim(cell1.Value) <> "" And cell1.EntireRow.Hidden = False And cell1.EntireColumn.Hidden = False Then
            Dim noMatch As Boolean
            noMatch = True
            For Each cell2 In rng2
                If cell2.Interior.Color <> RGB(0, 255, 0) And Trim(cell2.Value) <> "" And cell2.EntireRow.Hidden = False And cell2.EntireColumn.Hidden = False Then
                    If Levenshtein(cell1.Value, cell2.Value) <= threshold Then
                        noMatch = False
                        ' Add a comment with file names and matched values
                        cell1.AddComment Text:="Имя файла 1: " & fileName1 & vbCrLf & "Имя файла 2: " & fileName2 & vbCrLf & "Различия: " & vbCrLf & cell1.Value & vbCrLf & cell2.Value
                        ' Set comment size
                        cell1.Comment.Shape.Width = 500
                        cell1.Comment.Shape.Height = 100
                        Exit For
                    End If
                End If
            Next cell2

            If noMatch Then
                cell1.Interior.Color = RGB(255, 0, 0) ' Red
            Else
                cell1.Interior.Color = RGB(255, 255, 0) ' Yellow
            End If
        End If
    Next cell1
End Sub

Sub CallDataReconciliation(control As IRibbonControl)
    Call DataReconciliation
End Sub
