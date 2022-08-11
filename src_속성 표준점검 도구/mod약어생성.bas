Attribute VB_Name = "mod약어생성"
Option Explicit

Public Function 약어생성(aFullName As String, Optional aMaxSize = 4) As String
    Dim sFullName As String, sAbbName As String
    Dim i As Integer
    
    '영문 Full Name 중에 포함된 '-' 는 공백으로 변경하여 처리한다.
    sFullName = Trim(Replace(aFullName, "-", " "))
    
    '/* 영문 Full Name이 숫자로 시작하는 경우 단어 순서를 변경한다. (첫번째, 두번째 단어 순서만 변경) */
    Select Case Mid(sFullName, 1, 1)
        Case "0" To "9"
            sFullName = Trim(단어순서변경(sFullName))
    End Select

    sFullName = StrConv(sFullName, vbProperCase)
'    '/* 영문 Full Name 앞에 'A', 'An', 'The' 등의 관사는 약어생성시에 제거한다. */
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

    If InStr(1, sFullName, " ") = 0 Then '/* 영문FullName이 하나의 단어일 경우*/
        If Len(sFullName) <= 4 Then '/* 영문FullName의 길이가 4 이하 일 경우 */
            sAbbName = sFullName
        Else
            '/* 영문FullName의 4자 이후 단어들이 모두 모음인 경우 - 앞 4자리 단어를 축약어로 사용 */
            If 모음제거(Mid(sFullName, 5)) = "" Then
                sAbbName = Mid(sFullName, 1, 4)
            Else
                sAbbName = Mid(sFullName, 1, 1) + 모음제거(Mid(sFullName, 2))   '-- 모음시작 단어 첫모음 살림
                '/* 축약 결과 길이제한 : 최대 aMaxSize 자리까지 */
                If Len(sAbbName) > aMaxSize Then
                    sAbbName = Mid(sAbbName, 1, aMaxSize)
                End If
            End If

        End If
    Else '/* 영문FullName이 여러 단어로 구성되어 있을 경우 */
        Dim iBlankCnt As Integer
        iBlankCnt = Len(sFullName) - Len(Replace(sFullName, " ", ""))

        '각 단어 첫글자
        'sAbbName = WorksheetFunction.Proper(sFullName)
        'sAbbName = StrConv(sFullName, vbProperCase)
        sAbbName = sFullName

        Dim iPos As Integer, sChar As String, sAbbTmp As String
        If iBlankCnt = 1 Then '/* 단어수가 2개인 경우 - 각 단어의 시작 2글자 조합으로 생성 */
            sAbbName = 모음제거(sAbbName)
            i = InStr(1, sAbbName, " ") + 1
            sAbbName = UCase(Left(sAbbName, 2) + Mid(sAbbName, i, 2))
        ElseIf iBlankCnt <= 4 Then '/* 단어수가 4개 이하인 경우 - 첫글자 조합으로 생성 */
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
            '/* 단어의 길이가 최대 허용자리수 이하인 경우 - 그대로 사용 */
            If Len(Replace(sFullName, " ", "")) <= aMaxSize Then
                sAbbName = UCase(Replace(sFullName, " ", ""))
            Else
                Dim sInText As String
                sInText = 등위접속사정리(UCase(sFullName))
                '/* 등위접속사가 존재하면 첫글자들로 구성 */
                If sInText <> sFullName Then
                    sAbbName = 모음제거(sInText)
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
                    '/* 두단어로 구성된 경우, 각각의 단어에서 앞 2자리 문자를 가지고 약어 생성 */
                    Dim iUnitLen As Integer, aWordArray() As String, sWord As String
                    iUnitLen = 2
                    aWordArray = Split(sFullName)
                    For iPos = 0 To UBound(aWordArray)
                        sWord = aWordArray(iPos)
                        If Len(sWord) = iUnitLen Then
                            sAbbTmp = sAbbTmp + sWord
                        Else
                            If 모음제거(Mid(sWord, iUnitLen + 1)) = "" Then
                                sAbbTmp = sAbbTmp + Mid(sWord, 1, iUnitLen)
                            Else
                                sAbbTmp = sAbbTmp + Mid(Mid(sWord, 1, 1) + 모음제거(이중자음정리(Mid(sWord, 2))), 1, iUnitLen)
                            End If
                        End If
                    Next
                    sAbbName = sAbbTmp
                End If
            End If

        End If

    End If
    약어생성 = UCase(sAbbName)
End Function

Public Function 단어순서변경(aFullName As String) As String
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
    단어순서변경 = sResult
End Function

Public Function 모음제거(aFullName As String) As String
    모음제거 = Replace(aFullName, "A", "", , , vbTextCompare)
    모음제거 = Replace(모음제거, "E", "", , , vbTextCompare)
    모음제거 = Replace(모음제거, "I", "", , , vbTextCompare)
    모음제거 = Replace(모음제거, "O", "", , , vbTextCompare)
    모음제거 = Replace(모음제거, "U", "", , , vbTextCompare)
End Function

Public Function 이중자음정리(aFullName As String) As String
    이중자음정리 = Replace(aFullName, "BB", "B")
    이중자음정리 = Replace(이중자음정리, "CC", "C")
    이중자음정리 = Replace(이중자음정리, "DD", "D")
    이중자음정리 = Replace(이중자음정리, "FF", "F")
    이중자음정리 = Replace(이중자음정리, "GG", "G")
    이중자음정리 = Replace(이중자음정리, "HH", "H")
    이중자음정리 = Replace(이중자음정리, "JJ", "J")
    이중자음정리 = Replace(이중자음정리, "KK", "K")
    이중자음정리 = Replace(이중자음정리, "LL", "L")
    이중자음정리 = Replace(이중자음정리, "MM", "M")
    이중자음정리 = Replace(이중자음정리, "NN", "N")
    이중자음정리 = Replace(이중자음정리, "PP", "P")
    이중자음정리 = Replace(이중자음정리, "QQ", "Q")
    이중자음정리 = Replace(이중자음정리, "RR", "R")
    이중자음정리 = Replace(이중자음정리, "SS", "S")
    이중자음정리 = Replace(이중자음정리, "TT", "T")
    이중자음정리 = Replace(이중자음정리, "VV", "V")
    이중자음정리 = Replace(이중자음정리, "WW", "W")
    이중자음정리 = Replace(이중자음정리, "XX", "X")
    이중자음정리 = Replace(이중자음정리, "YY", "Y")
    이중자음정리 = Replace(이중자음정리, "ZZ", "Z")
End Function

'AND, OR, OF, BY 등의 등위접속사 정리
Public Function 등위접속사정리(aFullName As String) As String
    등위접속사정리 = Replace(aFullName, " AND ", "N")
    등위접속사정리 = Replace(등위접속사정리, " OR ", "R")
    등위접속사정리 = Replace(등위접속사정리, " OF ", "F")
    등위접속사정리 = Replace(등위접속사정리, " BY ", "B")
End Function


