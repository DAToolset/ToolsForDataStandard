Attribute VB_Name = "mod������"
Option Explicit

Public Function ������(aFullName As String, Optional aMaxSize = 4) As String
    Dim sFullName As String, sAbbName As String
    Dim i As Integer
    
    '���� Full Name �߿� ���Ե� '-' �� �������� �����Ͽ� ó���Ѵ�.
    sFullName = Trim(Replace(aFullName, "-", " "))
    
    '/* ���� Full Name�� ���ڷ� �����ϴ� ��� �ܾ� ������ �����Ѵ�. (ù��°, �ι�° �ܾ� ������ ����) */
    Select Case Mid(sFullName, 1, 1)
        Case "0" To "9"
            sFullName = Trim(�ܾ��������(sFullName))
    End Select

    sFullName = StrConv(sFullName, vbProperCase)
'    '/* ���� Full Name �տ� 'A', 'An', 'The' ���� ����� �������ÿ� �����Ѵ�. */
'    If Left(sFullName, 2) = "A " Then
'        sFullName = Mid(sFullName, 3)
'    ElseIf Left(sFullName, 3) = "AN " Then
'        sFullName = Mid(sFullName, 4)
'    ElseIf Left(sFullName, 4) = "THE " Then
'        sFullName = Mid(sFullName, 5)
'    End If
    sFullName = Replace(sFullName, "A ", "")
    sFullName = Replace(sFullName, "An ", "")
    sFullName = Replace(sFullName, "The ", "")

    If InStr(1, sFullName, " ") = 0 Then '/* ����FullName�� �ϳ��� �ܾ��� ���*/
        If Len(sFullName) <= 4 Then '/* ����FullName�� ���̰� 4 ���� �� ��� */
            sAbbName = sFullName
        Else
            '/* ����FullName�� 4�� ���� �ܾ���� ��� ������ ��� - �� 4�ڸ� �ܾ ����� ��� */
            If ��������(Mid(sFullName, 5)) = "" Then
                sAbbName = Mid(sFullName, 1, 4)
            Else
                sAbbName = Mid(sFullName, 1, 1) + ��������(Mid(sFullName, 2))   '-- �������� �ܾ� ù���� �츲
                '/* ��� ��� �������� : �ִ� aMaxSize �ڸ����� */
                If Len(sAbbName) > aMaxSize Then
                    sAbbName = Mid(sAbbName, 1, aMaxSize)
                End If
            End If

        End If
    Else '/* ����FullName�� ���� �ܾ�� �����Ǿ� ���� ��� */
        Dim iBlankCnt As Integer
        iBlankCnt = Len(sFullName) - Len(Replace(sFullName, " ", ""))

        '�� �ܾ� ù����
        'sAbbName = WorksheetFunction.Proper(sFullName)
        'sAbbName = StrConv(sFullName, vbProperCase)
        sAbbName = sFullName

        Dim iPos As Integer, sChar As String, sAbbTmp As String
        If iBlankCnt = 1 Then '/* �ܾ���� 2���� ��� - �� �ܾ��� ���� 2���� �������� ���� */
            sAbbName = ��������(sAbbName)
            i = InStr(1, sAbbName, " ") + 1
            sAbbName = UCase(Left(sAbbName, 2) + Mid(sAbbName, i, 2))
        ElseIf iBlankCnt <= 4 Then '/* �ܾ���� 4�� ������ ��� - ù���� �������� ���� */
             For iPos = 1 To Len(sAbbName)
                sChar = Mid(sAbbName, iPos, 1)
                If (sChar >= "A" And sChar <= "Z") Or (sChar >= "0" And sChar <= "9") Then
                    sAbbTmp = sAbbTmp + sChar
                End If
                If Len(sAbbTmp) >= aMaxSize Then GoTo Exit_For1
             Next iPos
Exit_For1:
             sAbbName = sAbbTmp
        Else
            '/* �ܾ��� ���̰� �ִ� ����ڸ��� ������ ��� - �״�� ��� */
            If Len(Replace(sFullName, " ", "")) <= aMaxSize Then
                sAbbName = UCase(Replace(sFullName, " ", ""))
            Else
                Dim sInText As String
                sInText = �������ӻ�����(UCase(sFullName))
                '/* �������ӻ簡 �����ϸ� ù���ڵ�� ���� */
                If sInText <> sFullName Then
                    sAbbName = ��������(sInText)
                    For iPos = 1 To Len(sAbbName)
                        sChar = Mid(sAbbName, iPos, 1)
                        If (sChar >= "A" And sChar <= "Z") Or (sChar >= "0" And sChar <= "9") Then
                            sAbbTmp = sAbbTmp + sChar
                        End If
                        If Len(sAbbTmp) >= aMaxSize Then GoTo Exit_For2
                    Next iPos
Exit_For2:
                    sAbbName = sAbbTmp
                Else
                    '/* �δܾ�� ������ ���, ������ �ܾ�� �� 2�ڸ� ���ڸ� ������ ��� ���� */
                    Dim iUnitLen As Integer, aWordArray() As String, sWord As String
                    iUnitLen = 2
                    aWordArray = Split(sFullName)
                    For iPos = 0 To UBound(aWordArray)
                        sWord = aWordArray(iPos)
                        If Len(sWord) = iUnitLen Then
                            sAbbTmp = sAbbTmp + sWord
                        Else
                            If ��������(Mid(sWord, iUnitLen + 1)) = "" Then
                                sAbbTmp = sAbbTmp + Mid(sWord, 1, iUnitLen)
                            Else
                                sAbbTmp = sAbbTmp + Mid(Mid(sWord, 1, 1) + ��������(������������(Mid(sWord, 2))), 1, iUnitLen)
                            End If
                        End If
                    Next
                    sAbbName = sAbbTmp
                End If
            End If

        End If

    End If
    ������ = UCase(sAbbName)
End Function

Public Function �ܾ��������(aFullName As String) As String
    Dim iFromIdx1 As Integer, iToIdx1 As Integer, sWord1 As String
    Dim iFromIdx2 As Integer, iToIdx2 As Integer, sWord2 As String
    Dim sResult As String, sWordRemain As String

    sResult = aFullName
    iFromIdx1 = 1
    iToIdx1 = InStr(1, aFullName, " ", vbTextCompare)
    If iToIdx1 = 0 Or iToIdx1 = Null Then GoTo Exit_Func
    sWord1 = Mid(aFullName, iFromIdx1, iToIdx1 - 1)

    iFromIdx2 = iToIdx1 + 1
    iToIdx2 = InStr(iFromIdx2, aFullName, " ", vbTextCompare)
    If iFromIdx2 = 0 Or iFromIdx2 = Null Then GoTo Exit_Func
    sWord2 = Mid(aFullName, iFromIdx2, iToIdx2 - iFromIdx2)

    sWordRemain = Mid(aFullName, iToIdx2 + 1, Len(aFullName) - iToIdx2)
    sResult = sWord2 + " " + sWord1 + " " + sWordRemain

Exit_Func:
    �ܾ�������� = sResult
End Function

Public Function ��������(aFullName As String) As String
    �������� = Replace(aFullName, "A", "", , , vbTextCompare)
    �������� = Replace(��������, "E", "", , , vbTextCompare)
    �������� = Replace(��������, "I", "", , , vbTextCompare)
    �������� = Replace(��������, "O", "", , , vbTextCompare)
    �������� = Replace(��������, "U", "", , , vbTextCompare)
End Function

Public Function ������������(aFullName As String) As String
    ������������ = Replace(aFullName, "BB", "B")
    ������������ = Replace(������������, "CC", "C")
    ������������ = Replace(������������, "DD", "D")
    ������������ = Replace(������������, "FF", "F")
    ������������ = Replace(������������, "GG", "G")
    ������������ = Replace(������������, "HH", "H")
    ������������ = Replace(������������, "JJ", "J")
    ������������ = Replace(������������, "KK", "K")
    ������������ = Replace(������������, "LL", "L")
    ������������ = Replace(������������, "MM", "M")
    ������������ = Replace(������������, "NN", "N")
    ������������ = Replace(������������, "PP", "P")
    ������������ = Replace(������������, "QQ", "Q")
    ������������ = Replace(������������, "RR", "R")
    ������������ = Replace(������������, "SS", "S")
    ������������ = Replace(������������, "TT", "T")
    ������������ = Replace(������������, "VV", "V")
    ������������ = Replace(������������, "WW", "W")
    ������������ = Replace(������������, "XX", "X")
    ������������ = Replace(������������, "YY", "Y")
    ������������ = Replace(������������, "ZZ", "Z")
End Function

'AND, OR, OF, BY ���� �������ӻ� ����
Public Function �������ӻ�����(aFullName As String) As String
    �������ӻ����� = Replace(aFullName, " AND ", "N")
    �������ӻ����� = Replace(�������ӻ�����, " OR ", "R")
    �������ӻ����� = Replace(�������ӻ�����, " OF ", "F")
    �������ӻ����� = Replace(�������ӻ�����, " BY ", "B")
End Function


