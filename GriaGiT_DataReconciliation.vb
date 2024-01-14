' Levenshtein Distance Function
' This function calculates the Levenshtein distance between two strings, which is the minimum number of single-character edits (insertions, deletions, or substitutions) required to change one string into the other.
Function Levenshtein(s1 As String, s2 As String) As Integer
    Dim i As Integer, j As Integer
    Dim l1 As Integer, l2 As Integer
    Dim dist() As Integer
    Dim min1 As Integer, min2 As Integer

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
                min1 = dist(i - 1, j) + 1
                min2 = dist(i, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                min2 = dist(i - 1, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                dist(i, j) = min1
            End If
        Next
    Next

    Levenshtein = dist(l1, l2)
End Function


' Data Reconciliation Subroutine
' This subroutine performs data reconciliation between two ranges in different workbooks. It first identifies exact matches and marks them with green color. Then, it compares non-matching cells based on the Levenshtein distance and marks them with red or yellow color depending on whether they are within the specified threshold.
Sub DataReconciliation()
    Dim rng1 As Range, rng2 As Range, cell1 As Range, cell2 As Range
    Dim word As Variant, words1 As Variant, words2 As Variant
    Dim exactMismatch As Boolean
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

    ' Clear existing formats and comments
    rng1.ClearFormats
    rng1.ClearComments
    'rng2.ClearFormats
    'rng2.ClearComments

    ' Create copies of ranges
    Dim rng1Copy As Range, rng2Copy As Range
    Set rng1Copy = rng1
    Set rng2Copy = rng2

    ' Sort copies of data in ascending order
    rng1Copy.Sort Key1:=rng1Copy, Order1:=xlAscending, Header:=xlNo
    rng2Copy.Sort Key1:=rng2Copy, Order1:=xlAscending, Header:=xlNo

    ' Ask user for the threshold value
    Dim threshold As Integer
    threshold = Application.InputBox("Введите допустимое количество различий", Type:=1)

    ' First pass: mark exact matches with green color
For Each cell1 In rng1
    For Each cell2 In rng2
        If Trim(cell1.Value) <> "" And Trim(cell2.Value) <> "" And Replace(cell1.Value, " ", "") = Replace(cell2.Value, " ", "") Then
            cell1.Interior.Color = RGB(0, 255, 0) ' Green
            Exit For
        End If
    Next cell2
Next cell1

' Second pass: compare non-matching cells based on Levenshtein distance
For Each cell1 In rng1
    ' Ignore matched cells and empty cells
    If cell1.Interior.Color <> RGB(0, 255, 0) And Trim(cell1.Value) <> "" Then
        Dim noMatch As Boolean
        noMatch = True
        For Each cell2 In rng2
            If cell2.Interior.Color <> RGB(0, 255, 0) And Trim(cell2.Value) <> "" Then ' Ignore matched cells
                If Levenshtein(cell1.Value, cell2.Value) <= threshold Then
                    noMatch = False
                    ' Add a comment with file names and matched values
                    cell1.AddComment Text:="Имя файла 1: " & fileName1 & vbCrLf & _
                                        "Имя файла 2: " & fileName2 & vbCrLf & _
                                        vbCrLf & "Различия: " & vbCrLf & cell1.Value & vbCrLf & cell2.Value
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

' Clear copies of ranges
Set rng1Copy = Nothing
Set rng2Copy = Nothing

End Sub

Sub CallDataReconciliation(control As IRibbonControl)
    Call DataReconciliation
End Sub

