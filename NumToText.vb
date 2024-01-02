Option Explicit
Public Function NumToText(chislo As Double, opc As Long) As String
    '
    ' opc=1 ������ ����� ��������
    ' opc=2 ����� ��������, ������� ����� ������
    ' opc=3 ����� �������� � ������� ����� ��������
 
    Dim drobLen As Variant, x2 As Double, utr As String, utr1, ITOG0 As String, itog As String
    Dim tekst As String, sklon As String, drobprop As String
    If chislo > 999999999999.99 Then
        NumToText = "�������� ������ 999 999 999 999.99!"
    ElseIf chislo < 0 Then
        NumToText = "�������� �������������!"
    End If
    Select Case opc
        Case 1
            chislo = Int(chislo)
            ITOG0 = �����OPC(chislo, opc)
        Case 2
            x2 = CDbl(Mid(chislo, InStr(1, chislo, ",") + 1))
            utr = Str(chislo)
            utr1 = Split(utr, ".")
            drobLen = Len(utr1(1))
            itog = �����OPC(Int(chislo), opc)
            If Right(Int(chislo), 1) = "1" Then
                tekst = tekst & " ����� "
            Else
                tekst = tekst & " ����� "
            End If
            sklon = drobnaya(x2, drobLen)
            ITOG0 = itog & tekst & x2 & sklon
        Case 3
            x2 = CDbl(Mid(chislo, InStr(1, chislo, ",") + 1))
            utr = Str(chislo)
            utr1 = Split(utr, ".")
            drobLen = Len(utr1(1))    '- 2
            itog = �����OPC(Int(chislo), opc)
            drobprop = �����OPC(Int(x2), opc)
            If Right(Int(chislo), 1) = "1" Then
                tekst = tekst & " ����� "
            Else
                tekst = tekst & " ����� "
            End If
            sklon = drobnaya(x2, drobLen)
            ITOG0 = itog & tekst & drobprop & sklon
    End Select
    NumToText = ITOG0
End Function
 
Public Function �����OPC(x As Double, opc As Long) As String
    Dim y(1 To 4) As Integer, i1 As Byte
    Dim Text(1 To 4) As String, i2 As Byte, y1 As Byte, y2 As Byte, _
            y3 As Byte, Text0 As String, text1 As String, text2 As String, Text3 As String, _
            Text4 As String
    For i1 = 1 To 4
        x = Fix(x) / 1000
        y(i1) = (x - Fix(x)) * 1000
    Next
    For i2 = 1 To 4
        y1 = y(i2) Mod 10
        y2 = (y(i2) - y1) / 10 Mod 10
        y3 = y(i2) \ 100
        text1 = Choose(y3 + 1, "", "��� ", "������ ", "������ ", "��������� ", _
                "������� ", "�������� ", "������� ", "��������� ", "��������� ")
        text2 = Choose(y2 + 1, "", "", "�������� ", "�������� ", "����� ", _
                "��������� ", "���������� ", "��������� ", "����������� ", "��������� ")
        If y2 = 1 Then
            Text3 = Choose(y1 + 1, "������ ", "����������� ", "���������� ", _
                    "���������� ", "������������ ", "���������� ", "����������� ", _
                    "���������� ", "������������ ", "������������ ")
        ElseIf y2 <> 1 And i2 = 2 Then
            Text3 = Choose(y1 + 1, "", "���� ", "��� ", "��� ", "������ ", "���� ", _
                    "����� ", "���� ", "������ ", "������ ")
        Else
            If opc = 2 Or opc = 3 Then
                Text3 = Choose(y1 + 1, "", "���� ", "��� ", "��� ", "������ ", "���� ", _
                        "����� ", "���� ", "������ ", "������ ")
            Else
                Text3 = Choose(y1 + 1, "", "���� ", "��� ", "��� ", "������ ", "���� ", _
                        "����� ", "���� ", "������ ", "������ ")
            End If
        End If
        If y2 <> 1 And y1 = 1 Then
            Text4 = Choose(i2, "", "������ ", "������� ", "�������� ")
        ElseIf y2 <> 1 And y1 > 1 And y1 < 5 Then
            Text4 = Choose(i2, "", "������ ", "�������� ", "��������� ")
        ElseIf y1 = 0 And y2 = 0 And y3 = 0 Then
            Text4 = Choose(i2, "", "", "", "")
        Else
            Text4 = Choose(i2, "", "����� ", "��������� ", "���������� ")
        End If
        Text(i2) = text1 & text2 & Text3 & Text4
    Next
    If y(1) + y(2) + y(3) + y(4) = 0 Then
        Text0 = "���� "
    Else
        Text0 = Text(4) & Text(3) & Text(2) & Text(1)
    End If
    �����OPC = Text0
End Function
 
Public Function drobnaya(x2 As Double, drobLen As Variant) As String
    Dim x As Variant, scl As String
    x = Right(x2, 1)
    If x = 1 Then
        scl = Choose(drobLen, "�������", "�����", "��������", "��������������", _
                "�����������", "����������", "����������������", "�������������", _
                "�����������", "�����������������")
    Else
        scl = Choose(drobLen, "�������", "�����", "��������", "��������������", _
                "�����������", "����������", "����������������", "�������������", _
                "�����������", "�����������������")
    End If
    drobnaya = scl
End Function

Sub InsertFormula()
    Dim �������������� As Range
    Dim ����� As Integer
    Dim ������� As String
    
    ' ����� ������ ��� ������� ��������
    On Error Resume Next ' ��������� ������, ���� �� ������� ������
    Set �������������� = Application.InputBox("�������� ������ ��� ������� ��������:", Type:=8)
    On Error GoTo 0 ' ���������� ����������� ���������� ������
    
    If �������������� Is Nothing Then Exit Sub ' ���� �� ������� ������, ������� �� �������
    
    ' ����� ������ ������ �������
    ����� = Application.InputBox("�������� ����� ������ ������� (1, 2 ��� 3):", Type:=1)
    
    If ����� < 1 Or ����� > 3 Then Exit Sub ' ���� ����� �� ������ ��� ������ �������, ������� �� �������
    
    ' ������������ ������ �������
    ������� = "=NumToText(" & ��������������.Address & "," & ����� & ")"
    
    ' ������� ������� � �������� ������
    ActiveCell.Formula = �������
End Sub


Sub CallInsertFormula(control As IRibbonControl)
    Call InsertFormula
End Sub