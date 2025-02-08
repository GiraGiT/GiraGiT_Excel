' ������� ������� � ������������ ������
Function CleanString(ByVal str As String) As String
    ' �������� ������ � �������� ��������
    str = UCase(str)
    
    ' ������� ��� ������� � ����� ����������
    Dim i As Integer
    Dim result As String
    For i = 1 To Len(str)
        Dim ch As String
        ch = Mid(str, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "�" And ch <= "�") Or (ch >= "0" And ch <= "9") Then
            result = result & ch
        End If
    Next i
    str = result
    
    ' ������ �������� �� �������� ��������
    Dim replacements As Variant
    replacements = Array( _
        Array("0", "O"), Array("1", "I"), Array("3", "E"), Array("4", "A"), _
        Array("5", "S"), Array("6", "G"), Array("7", "T"), Array("8", "B"), _
        Array("9", "G"), Array("2", "Z"), _
        Array("C", "�"), Array("c", "�"), _
        Array("�������", "������"), _
        Array("�", "A"), Array("�", "B"), Array("�", "E"), Array("�", "K"), Array("�", "M"), _
        Array("�", "H"), Array("�", "O"), Array("�", "P"), Array("�", "C"), Array("�", "T"), _
        Array("�", "Y"), Array("�", "X"), Array("�", ""), Array("�", "I"), Array("�", "E"), _
        Array("�", "I"), Array("�", "J"), Array("�", "D"), Array("�", "L"), Array("�", "F"), _
        Array("�", "Z"), Array("�", "C"), Array("�", "CH"), Array("�", "SH"), Array("�", "SCH"), _
        Array("�", "G"), Array("�", "P"), Array("�", "ZH"), Array("�", "YU"), Array("�", "YA"), _
        Array("�", "B"), Array("�", ""), Array("�", ""), _
        Array("A", "�"), Array("B", "�"), Array("E", "�"), Array("K", "�"), Array("M", "�"), _
        Array("H", "�"), Array("O", "�"), Array("P", "�"), Array("C", "�"), Array("T", "�"), _
        Array("Y", "�"), Array("X", "�"), Array("I", "�"), Array("J", "�"), Array("G", "�"), _
        Array("L", "�"), Array("D", "�"), Array("F", "�"), Array("Z", "�"), Array("N", "�"), _
        Array("Q", "�"), Array("S", "�"), Array("V", "�"), Array("U", "�"), Array("W", "�"))
    
    str = ReplaceCharacters(str, replacements)
    
    CleanString = str
End Function

' ������� ������ �������� �� �������� ��������
Function ReplaceCharacters(ByVal str As String, ByVal replacements As Variant) As String
    Dim i As Integer
    For i = 0 To UBound(replacements)
        str = Replace(str, replacements(i)(0), replacements(i)(1))
    Next i
    ReplaceCharacters = str
End Function

' ������� ������ ���������� ����� ���������
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

' ������� ���������� �������� ����������
Function SubstringSimilarity(str1 As String, str2 As String) As Double
    Dim s1 As String, s2 As String
    Dim commonSub As String
    Dim perc1 As Double, perc2 As Double
    
    s1 = CleanString(str1)
    s2 = CleanString(str2)

    ' ������� ���������� ����������
    Debug.Print "��������1: " & str1 & " | ���������1: " & s1
    Debug.Print "��������2: " & str2 & " | ���������2: " & s2
    
    If Len(s1) = 0 Or Len(s2) = 0 Then
        SubstringSimilarity = 0
        Exit Function
    End If
    
    commonSub = LongestCommonSubstring(s1, s2)
    perc1 = Len(commonSub) / Len(s1)
    perc2 = Len(commonSub) / Len(s2)
    SubstringSimilarity = (perc1 + perc2) / 2
    
    ' ������� ������� ��������
    Debug.Print "��������: " & SubstringSimilarity
End Function

' �������� ��������� ��������� � ����������� ����� � ����������� ������������
Sub DataReconciliation()
    Dim firstRange As Range, secondRange As Range
    Dim cell As Range, matchCell As Range
    Dim bestSim As Double, sim As Double
    Dim bestMatch As String, bestComment As String
    Dim threshold As Double
    Dim commentText As String
    Dim fileName1 As String, sheetName1 As String, fileName2 As String, sheetName2 As String
    
    threshold = 0.5  ' �������� ����� �������� � 0.7 �� 0.5
    
    On Error Resume Next
    Set firstRange = Application.InputBox("�������� ������ �������� �����:", Type:=8)
    If firstRange Is Nothing Then Exit Sub
    Set secondRange = Application.InputBox("�������� ������ �������� �����:", Type:=8)
    If secondRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' �������� ���������� � ������ � ������
    fileName1 = firstRange.Worksheet.Parent.Name
    sheetName1 = firstRange.Worksheet.Name
    fileName2 = secondRange.Worksheet.Parent.Name
    sheetName2 = secondRange.Worksheet.Name
    
    ' ������� ���������� �������������� � �����������
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
                bestComment = "��� ����� ������� ���������: " & fileName1 & " > ����: " & sheetName1 & vbCrLf & _
                              "��� ����� ������� ���������: " & fileName2 & " > ����: " & sheetName2 & vbCrLf & _
                              "��������:" & vbCrLf & _
                              "������ ��������: " & cell.Text & vbCrLf & _
                              "������ ��������: " & matchCell.Text
            End If
NextMatch:
        Next matchCell
        
        ' ���������� � ��������� ����������� ��� �������������
        If bestSim = 1 Then
            cell.Interior.color = RGB(0, 255, 0)  ' ������
        ElseIf bestSim >= threshold Then
            cell.Interior.color = RGB(255, 255, 0)  ' Ƹ����
            commentText = fileName1 & " > " & sheetName1 & vbCrLf & _
                          fileName2 & " > " & sheetName2 & vbCrLf & vbCrLf & _
                          "��������:" & vbCrLf & _
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
            cell.Interior.color = RGB(255, 0, 0)  ' �������
        End If
NextCell:
    Next cell
    
    ' �������� ������� ��������� �������� � ������������
    HighlightDifferences
End Sub

' ������� ��������� ��������
Sub HighlightDifferences()
    Dim ws As Worksheet
    Dim cell As Range
    Dim commentText As String
    Dim i As Integer
    
    ' ������������ ��� ����������� �� �����
    For Each ws In ThisWorkbook.Worksheets
        For Each cell In ws.UsedRange
            If Not cell.Comment Is Nothing Then
                commentText = cell.Comment.Text
                
                ' ������������ ������ ������ �����������
                For i = 1 To Len(commentText)
                    ' ���������� ������ "G" ��� "C" � ������� ����
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