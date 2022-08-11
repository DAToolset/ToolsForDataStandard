Attribute VB_Name = "modControlStd"
Option Explicit

Public Const C_VERSION_STRING As String = "v1.20"

Public Enum StdDicMatchOption
    WordAndTerm = 1
    WordOnly = 2
    TermOnly = 3
End Enum

Public Enum WordMatchDirection
    LtoR = 1
    RtoL = 2
    Both = 3
End Enum

Dim oStdWordDic As CStdWordDic '단어사전
Dim oStdTermDic As CStdTermDic '용어사전
Dim oStdDomainDic As CStdDomainDic '도메인사전

Dim eStdDicMatchOption As StdDicMatchOption '표준사전 찾기 옵션
Dim eWordMatchDirection As WordMatchDirection '단어조합 방향

'표준점검 실행 프로시져
'* Parameter
'  - aAttrBaseRange              : 입력-점검대상 속성 목록의 시작 위치
'  - aLNameBaseRange             : 출력-표준단어 논리명 조합의 시작 위치
'  - aPNameBaseRange             : 출력-표준단어 물리명 조합의 시작 위치
'  - aStdWordBaseRange           : 사전-표준단어사전의 시작 위치
'  - aStdTermBaseRange           : 사전-표준용어사전의 시작 위치
'  - aStdDomainBaseRange         : 사전-표준도메인사전의 시작 위치
'  - aStdDicMatchOption          : 옵션-표준사전찾기 옵션 선택 값(1:단어&용어, 2:단어, 3:용어)
'  - aWordMatchDirection         : 옵션-단어조합방향 옵션 선택 값(1:좌->우, 2:우->좌, 3:모두)
'  - aIsAllowDupWordLogicalName  : 옵션-표준단어 논리명 중복(동음이의어) 허용 여부
'  - aIsAllowDupWordPhysicalName : 옵션-표준단어 물리명 중복(이음동의어) 허용 여부
'  - aIsOnlyForSelectedAttr      : 옵션-선택한 속성만 점검할지 여부(True: 선택한 속성만 점검, False: 전체 속성 점검)
'  - aSelectedAttrRange          : 옵션-선택한 속성명 목록의 범위
Public Sub 표준점검(aAttrBaseRange As Range, aLNameBaseRange As Range, aPNameBaseRange As Range, _
        aStdWordBaseRange As Range, aStdTermBaseRange As Range, aStdDomainBaseRange As Range, _
        aStdDicMatchOption As StdDicMatchOption, aWordMatchDirection As WordMatchDirection, _
        aIsAllowDupWordLogicalName As Boolean, aIsAllowDupWordPhysicalName As Boolean, _
        aIsOnlyForSelectedAttr As Boolean, _
        aSelectedAttrRange As Range)

    Application.ScreenUpdating = False
    eStdDicMatchOption = aStdDicMatchOption
    eWordMatchDirection = aWordMatchDirection
    'GetDicMap
    '단어사전 불러오기 ---------------------------------------------------------------------
    Set oStdWordDic = New CStdWordDic: oStdWordDic.m_eStdDicMatchOption = aStdDicMatchOption
    oStdWordDic.Load aStdWordBaseRange, aIsAllowDupWordLogicalName, aIsAllowDupWordPhysicalName
    If oStdWordDic.Count = 0 Then
        MsgBox "단어사전이 비어 있습니다." + vbLf + "표준 점검을 중지합니다", vbCritical, "오류"
        Exit Sub
    End If

    '용어사전 불러오기 ---------------------------------------------------------------------
    Set oStdTermDic = New CStdTermDic: oStdTermDic.m_eStdDicMatchOption = aStdDicMatchOption
    oStdTermDic.SetStdWordDicP oStdWordDic.m_oStdWordDicP
    oStdTermDic.Load aStdTermBaseRange

    '도메인사전 불러오기 ---------------------------------------------------------------------
    Set oStdDomainDic = New CStdDomainDic: oStdDomainDic.m_eStdDicMatchOption = aStdDicMatchOption
    oStdDomainDic.SetStdWordDic oStdWordDic
    oStdDomainDic.Load aStdDomainBaseRange

    Dim oAttrRange As Range, lRowOffset As Long, iFIdx As Integer
    Dim sAttrName As String, iLenAttrName As Integer, sAttrDataTypeSize As String
    'Dim sKorParseResult As String, sEngParseResult As String
    Dim saKorParseResult() As String, saEngParseResult() As String
    'Dim sKorParseResultR As String, sEngParseResultR As String
    Dim saKorParseResultR() As String, saEngParseResultR() As String
    Dim lParseResultIdx As Long
    Dim oStdWord As CStdWord, oStdTerm As CStdTerm, oStdWordObj As Object
    Dim iTokenLen As Integer, sToken As String
    Dim oWordMatchCol As Collection
    Dim sGenType As String, sStdTermDataTypeSize As String
    Dim sSuffix As String, sLastWord As String, sLastWordChk As String
    Dim bWordMatched As Boolean, bTermMatched As Boolean, b동음이의어Matched As Boolean, b이음동의어Mached As Boolean
    Dim b비표준단어MatchedL As Boolean, b비표준단어MatchedR As Boolean
    Dim vInRngArr As Variant, vOutRngArr As Variant, lRow As Long, lLastRow As Long, oOutRange As Range
    Dim lParseResultOffset As Long, lTgtIdx As Long, lIdx As Long

    lRowOffset = 0
    '입력범위 읽어서 Variant array에 담기
    If aAttrBaseRange.Offset(1, 0).Value2 = "" Then '첫 행만 데이터가 있는 경우
        vInRngArr = aAttrBaseRange.Resize(, 2).Value2 '읽는 범위: 2개 컬럼
    Else
        If Not aIsOnlyForSelectedAttr Then '선택한 속성만 읽지 않는 경우 --> 전체 읽기
            vInRngArr = Range(aAttrBaseRange, aAttrBaseRange.End(xlDown)).Resize(, 2).Value2 '읽는 범위: 2개 컬럼
        Else '선택한 속성만 읽는 경우
            lRowOffset = aSelectedAttrRange.Row - aLNameBaseRange.Row
            vInRngArr = aSelectedAttrRange.Value2
        End If
    End If

    lLastRow = UBound(vInRngArr) 'aAttrBaseRange.End(xlDown).Row
    'Set oOutRange = aLNameBaseRange.Resize(lLastRow, 8) '쓰는 범위: 8개 컬럼
    Set oOutRange = aLNameBaseRange.Offset(lRowOffset, 0).Resize(lLastRow, 8) '쓰는 범위: 8개 컬럼
    'vOutRngArr = oOutRange.Value2
    ReDim vOutRngArr(1 To oOutRange.Rows.Count, 1 To oOutRange.Columns.Count)

    For lRow = LBound(vInRngArr) To UBound(vInRngArr)
        sAttrName = vInRngArr(lRow, 1)
        If Trim(sAttrName) = "" Then GoTo SkipBlank '점검할 속성명이 비어있는 경우 Skip
        sAttrDataTypeSize = vInRngArr(lRow, 2)
        bWordMatched = False: bTermMatched = False: lParseResultIdx = 0: b동음이의어Matched = False
        b비표준단어MatchedL = False: b비표준단어MatchedR = False: b이음동의어Mached = False

'        sKorParseResult = "": sEngParseResult = "": sStdTermDataTypeSize = ""
'        sKorParseResultR = "": sEngParseResultR = "": sSuffix = ""

        '------------------------------------------------------------------------------------------
        '변수 초기화
        ReDim saKorParseResult(1 To 1): ReDim saEngParseResult(1 To 1)
        ReDim saKorParseResultR(1 To 1): ReDim saEngParseResultR(1 To 1)
        sStdTermDataTypeSize = "": sSuffix = "": sGenType = "": sLastWord = "": sLastWordChk = ""
        '------------------------------------------------------------------------------------------

        If Not IsValidAttributeName(sAttrName) Then '속성명에 부적합한 문자가 포함된 경우 메시지 보여주고 Skip 또는 중단
            If IsOkToGo("속성명에 부적합한 문자가 포함되어 있습니다." + vbLf + _
                        "- 부적합한 문자: [, 행분리문자" + vbLf + vbLf + _
                        " - [Row#: " & CStr(lRow) + "] " & sAttrName + vbLf + vbLf + _
                        "해당 속성을 제외하고 계속 진행하시겠습니까?", "확인") Then
                saKorParseResult(1) = sAttrName
                sGenType = "부적합한 문자 확인필요"
                GoTo Continue_OuterFor1
            Else
                MsgBox "사용자의 요청으로 실행을 중지합니다.", vbOKOnly + vbCritical, "중지"
                GoTo Finalize
            End If
        End If

        '------------------------------------------------------------------------------------------
        '속성명 후위 숫자 제외
        sSuffix = GetNumberSuffix(sAttrName)
        If sSuffix <> "" Then sAttrName = GetTextWithoutSuffix(sAttrName, sSuffix)
        '------------------------------------------------------------------------------------------

        '------------------------------------------------------------------------------------------
        '속성명을 공백 또는 "_"로 구분한 경우 처리
        If (InStr(1, sAttrName, " ") > 0) Or (InStr(1, sAttrName, "_") > 0) Then
            Dim sAttrNameTemp As String, aAttrName() As String
            sAttrNameTemp = sAttrName
            sAttrNameTemp = Replace(sAttrNameTemp, "_", " ")
            aAttrName = Split(sAttrNameTemp, " ")

            Set oWordMatchCol = Nothing
            Set oWordMatchCol = New Collection
            For iFIdx = 0 To UBound(aAttrName)
                sToken = Trim(aAttrName(iFIdx))
                If sToken = "" Then GoTo Continue_InnerFor1
                If oStdWordDic.Exists(sToken) Then 'Token이 표준단어 목록에 포함된 경우
                    Set oStdWordObj = oStdWordDic.Item(sToken)
                    If TypeOf oStdWordObj Is CStdWord Then
                        '논리명이 유일한 경우
                        Set oStdWord = oStdWordObj
                        Set oStdWord = oStdWord.GetSWForNSW
                        If Not b비표준단어MatchedL Then
                            b비표준단어MatchedL = sToken <> oStdWord.m_s단어논리명
                        End If
                        If oStdWord.m_b물리명중복여부 = True Then b이음동의어Mached = True
                        saKorParseResult(1) = saKorParseResult(1) + IIf(saKorParseResult(1) = "", "", "_") + oStdWord.m_s단어논리명
                        For lParseResultIdx = LBound(saEngParseResult) To UBound(saEngParseResult) '동음이의어의 n개 물리명 조합
                            saEngParseResult(lParseResultIdx) = saEngParseResult(lParseResultIdx) + _
                                    IIf(saEngParseResult(lParseResultIdx) = "", "", "_") + oStdWord.m_s단어물리명
                        Next lParseResultIdx
                    ElseIf TypeOf oStdWordObj Is CStdWordCol Then
                        '동음이의어가 있어 논리명이 2개 이상인 경우
                        b동음이의어Matched = True
                        lParseResultOffset = UBound(saEngParseResult)
                        SetInitParseResult saEngParseResult, oStdWordObj.Count

                        Set oStdWord = oStdWordObj.Items(1)
                        saKorParseResult(1) = saKorParseResult(1) + IIf(saKorParseResult(1) = "", "", "_") + oStdWord.m_s단어논리명

                        lTgtIdx = 1
                        For lParseResultIdx = 1 To oStdWordObj.Count
                            Set oStdWord = oStdWordObj.Items(lParseResultIdx)
                            For lIdx = 1 To lParseResultOffset
                                saEngParseResult(lTgtIdx) = saEngParseResult(lTgtIdx) + IIf(saEngParseResult(lTgtIdx) = "", "", "_") + oStdWord.m_s단어물리명
                                lTgtIdx = lTgtIdx + 1
                            Next lIdx
                        Next lParseResultIdx
                    End If
                    bWordMatched = True
                Else '매칭된 단어가 없는 경우
                    sToken = IIf(saKorParseResult(1) = "", "[", "_[") + sToken + "]"
                    saKorParseResult(1) = saKorParseResult(1) + sToken
                    For lParseResultIdx = LBound(saEngParseResult) To UBound(saEngParseResult)
                        saEngParseResult(lParseResultIdx) = saEngParseResult(lParseResultIdx) + sToken
                    Next lParseResultIdx
                End If
Continue_InnerFor1:
            Next iFIdx
            sGenType = "표준단어 " + IIf(bWordMatched, "조합", "없음") + "(사용자 지정)" _
                     + IIf(b동음이의어Matched, vbLf + "(동음이의어 확인필요)", "") _
                     + IIf(b비표준단어MatchedL, vbLf + "(비표준단어 확인필요)", "")
            GoTo SkipIfTermMatched
        End If
        '------------------------------------------------------------------------------------------

        iLenAttrName = Len(sAttrName)

        '------------------------------------------------------------------------------------------
        '표준용어 찾기(단어로만 찾는 경우가 아닐 때)
        If (Not eStdDicMatchOption = WordOnly) And _
           (oStdTermDic.Exists(sAttrName)) Then  '속성명과 일치하는 표준용어 존재
            Set oStdTerm = oStdTermDic.Item(sAttrName)
            saKorParseResult(1) = oStdTerm.m_s단어논리명조합
            saEngParseResult(1) = oStdTerm.m_s용어물리명
            sStdTermDataTypeSize = oStdTerm.m_s데이터타입길이명
            sGenType = "표준용어 일치"
            bTermMatched = True
            GoTo SkipIfTermMatched
        End If
        '------------------------------------------------------------------------------------------

        If eStdDicMatchOption = TermOnly Then
            sGenType = "표준용어 없음"
            GoTo SkipIfTermMatched
        End If

        If eWordMatchDirection = RtoL Then GoTo Skip_LtoR
        '------------------------------------------------------------------------------------------
        '표준단어 조합하기 (좌 --> 우 탐색)
        iFIdx = 1
        Do
            If iFIdx > iLenAttrName Then Exit Do
            Set oWordMatchCol = Nothing
            Set oWordMatchCol = New Collection
            For iTokenLen = 1 To iLenAttrName - iFIdx + 1 'iMaxStdWordLen
                'Token 생성
                sToken = Mid(sAttrName, iFIdx, iTokenLen)
                If oStdWordDic.Exists(sToken) Then 'Token이 표준단어 목록에 포함된 경우
                    Set oStdWordObj = oStdWordDic.Item(sToken)
                    oWordMatchCol.Add oStdWordObj
                End If
            Next iTokenLen
            If oWordMatchCol.Count > 0 Then '매칭된 단어가 있는 경우
                Set oStdWordObj = oWordMatchCol(oWordMatchCol.Count) '가장 긴 표준단어 매칭결과 적용
                If TypeOf oStdWordObj Is CStdWord Then
                    '논리명이 유일한 경우
                    Set oStdWord = oStdWordObj
                    sToken = oStdWord.m_s단어논리명
                    Set oStdWord = oStdWord.GetSWForNSW
                    If Not b비표준단어MatchedL Then
                        b비표준단어MatchedL = sToken <> oStdWord.m_s단어논리명
                    End If
                    If oStdWord.m_b물리명중복여부 = True Then b이음동의어Mached = True
                    saKorParseResult(1) = saKorParseResult(1) + IIf(saKorParseResult(1) = "", "", "_") + oStdWord.m_s단어논리명
                    For lParseResultIdx = LBound(saEngParseResult) To UBound(saEngParseResult) '동음이의어의 n개 물리명 조합
                        saEngParseResult(lParseResultIdx) = saEngParseResult(lParseResultIdx) + _
                                IIf(saEngParseResult(lParseResultIdx) = "", "", "_") + oStdWord.m_s단어물리명
                    Next lParseResultIdx

                ElseIf TypeOf oStdWordObj Is CStdWordCol Then
                    '동음이의어가 있어 논리명이 2개 이상인 경우
                    b동음이의어Matched = True
                    lParseResultOffset = UBound(saEngParseResult)
                    SetInitParseResult saEngParseResult, oStdWordObj.Count

                    Set oStdWord = oStdWordObj.Items(1)
                    sToken = oStdWord.m_s단어논리명
                    Set oStdWord = oStdWord.GetSWForNSW
                    If Not b비표준단어MatchedL Then
                        b비표준단어MatchedL = sToken <> oStdWord.m_s단어논리명
                    End If
                    saKorParseResult(1) = saKorParseResult(1) + IIf(saKorParseResult(1) = "", "", "_") + oStdWord.m_s단어논리명

                    lTgtIdx = 1
                    For lParseResultIdx = 1 To oStdWordObj.Count
                        Set oStdWord = oStdWordObj.Items(lParseResultIdx)
                        Set oStdWord = oStdWord.GetSWForNSW
                        For lIdx = 1 To lParseResultOffset
                            saEngParseResult(lTgtIdx) = saEngParseResult(lTgtIdx) + IIf(saEngParseResult(lTgtIdx) = "", "", "_") + oStdWord.m_s단어물리명
                            lTgtIdx = lTgtIdx + 1
                        Next lIdx
                    Next lParseResultIdx
                End If

                'iFIdx = iFIdx + Len(oStdWord.m_s단어논리명)
                iFIdx = iFIdx + Len(sToken)
                bWordMatched = True
            Else '매칭된 단어가 없는 경우
                sToken = Mid(sAttrName, iFIdx, 1)
                sToken = IIf(saKorParseResult(1) = "", "[", "_[") + sToken + "]"
                saKorParseResult(1) = saKorParseResult(1) + sToken
                For lParseResultIdx = LBound(saEngParseResult) To UBound(saEngParseResult)
                    saEngParseResult(lParseResultIdx) = saEngParseResult(lParseResultIdx) + sToken
                Next lParseResultIdx
                iFIdx = iFIdx + 1
            End If
        Loop
        '------------------------------------------------------------------------------------------
Skip_LtoR:

        If eWordMatchDirection = LtoR Then GoTo Skip_RtoL
        '------------------------------------------------------------------------------------------
        '표준단어 조합하기 (우 --> 좌 탐색)
        iFIdx = iLenAttrName
        Do
            If iFIdx <= 0 Then Exit Do
            Set oWordMatchCol = Nothing
            Set oWordMatchCol = New Collection
            For iTokenLen = 1 To iFIdx
                'Token 생성
                sToken = Mid(sAttrName, iFIdx - iTokenLen + 1, iTokenLen)
                If oStdWordDic.Exists(sToken) Then 'Token이 표준단어 목록에 포함된 경우
                    Set oStdWordObj = oStdWordDic.Item(sToken)
                    oWordMatchCol.Add oStdWordObj
                End If
            Next iTokenLen
            If oWordMatchCol.Count > 0 Then '매칭된 단어가 있는 경우
                Set oStdWordObj = oWordMatchCol(oWordMatchCol.Count) '가장 긴 표준단어 매칭결과 적용
                If TypeOf oStdWordObj Is CStdWord Then
                    '논리명이 유일한 경우
                    Set oStdWord = oStdWordObj
                    sToken = oStdWord.m_s단어논리명
                    Set oStdWord = oStdWord.GetSWForNSW
                    If Not b비표준단어MatchedR Then
                        b비표준단어MatchedR = sToken <> oStdWord.m_s단어논리명
                    End If
                    If oStdWord.m_b물리명중복여부 = True Then b이음동의어Mached = True
                    iFIdx = iFIdx - Len(sToken)
                    saKorParseResultR(1) = IIf(iFIdx > 0, "_", "") + oStdWord.m_s단어논리명 + saKorParseResultR(1)
                    For lParseResultIdx = LBound(saEngParseResultR) To UBound(saEngParseResultR) '동음이의어의 n개 물리명 조합
                        saEngParseResultR(lParseResultIdx) = IIf(iFIdx > 0, "_", "") + oStdWord.m_s단어물리명 + saEngParseResultR(lParseResultIdx)
                    Next lParseResultIdx
                ElseIf TypeOf oStdWordObj Is CStdWordCol Then
                    '동음이의어가 있어 논리명이 2개 이상인 경우
                    b동음이의어Matched = True
                    lParseResultOffset = UBound(saEngParseResultR)
                    SetInitParseResult saEngParseResultR, oStdWordObj.Count

                    Set oStdWord = oStdWordObj.Items(1)
                    sToken = oStdWord.m_s단어논리명
                    Set oStdWord = oStdWord.GetSWForNSW
                    If Not b비표준단어MatchedR Then
                        b비표준단어MatchedR = sToken <> oStdWord.m_s단어논리명
                    End If
                    iFIdx = iFIdx - Len(sToken)
                    saKorParseResultR(1) = IIf(iFIdx > 0, "_", "") + oStdWord.m_s단어논리명 + saKorParseResultR(1)

                    lTgtIdx = 1
                    For lParseResultIdx = 1 To oStdWordObj.Count
                        Set oStdWord = oStdWordObj.Items(lParseResultIdx)
                        Set oStdWord = oStdWord.GetSWForNSW
                        For lIdx = 1 To lParseResultOffset
                            saEngParseResultR(lTgtIdx) = IIf(iFIdx > 0, "_", "") + oStdWord.m_s단어물리명 + saEngParseResultR(lTgtIdx)
                            lTgtIdx = lTgtIdx + 1
                        Next lIdx
                    Next lParseResultIdx
                End If
                
                bWordMatched = True
            Else '매칭된 단어가 없는 경우
                sToken = Mid(sAttrName, iFIdx, 1)
                sToken = IIf(iFIdx > 1, "_", "") + "[" + sToken + "]"
                saKorParseResultR(1) = sToken + saKorParseResultR(1)
                For lParseResultIdx = LBound(saEngParseResultR) To UBound(saEngParseResultR)
                    saEngParseResultR(lParseResultIdx) = sToken + saEngParseResultR(lParseResultIdx)
                Next lParseResultIdx
                iFIdx = iFIdx - 1
            End If
        Loop
        '------------------------------------------------------------------------------------------
Skip_RtoL:
        sGenType = "표준단어 " + IIf(bWordMatched, "조합", "없음") _
                 + IIf(b동음이의어Matched, vbLf + "(동음이의어 확인필요)", "") _
                 + IIf(b비표준단어MatchedL Or b비표준단어MatchedR, vbLf + "(비표준단어 확인필요)", "") _
                 + IIf(b이음동의어Mached, vbLf + "(이음동의어 확인필요)", "")
        If eWordMatchDirection = RtoL Then
            saKorParseResult(1) = saKorParseResultR(1)
            saEngParseResult(1) = saEngParseResultR(1)
        End If

        '좌->우 매칭과 우->좌 매칭의 결과가 다를 경우 모두 보여지게
        If (eWordMatchDirection = Both) And _
           (saKorParseResult(1) <> saKorParseResultR(1)) Then
            lParseResultOffset = UBound(saKorParseResult)
            ReDim Preserve saKorParseResult(1 To UBound(saKorParseResult) + UBound(saKorParseResultR))
            For lIdx = 1 To UBound(saKorParseResultR)
                saKorParseResult(lParseResultOffset + lIdx) = saKorParseResultR(lIdx)
            Next lIdx

            lParseResultOffset = UBound(saEngParseResult)
            ReDim Preserve saEngParseResult(1 To UBound(saEngParseResult) + UBound(saEngParseResultR))
            For lIdx = 1 To UBound(saEngParseResultR)
                saEngParseResult(lParseResultOffset + lIdx) = saEngParseResultR(lIdx)
            Next lIdx

            sGenType = sGenType + vbLf + "(조합 패턴 확인필요)"
        End If

SkipIfTermMatched:
Continue_OuterFor1:
        '표준단어 논리명 조합
        For lParseResultIdx = LBound(saKorParseResult) To UBound(saKorParseResult)
'            vOutRngArr(lRow, 1) = vOutRngArr(lRow, 1) + saKorParseResult(lParseResultIdx) + IIf(sSuffix = "", "", "_" + sSuffix) + _
'                                  IIf(lParseResultIdx = UBound(saKorParseResult), "", vbLf)
            vOutRngArr(lRow, 1) = vOutRngArr(lRow, 1) + saKorParseResult(lParseResultIdx) + sSuffix + _
                                  IIf(lParseResultIdx = UBound(saKorParseResult), "", vbLf)
        Next lParseResultIdx

        '표준단어 물리명 조합
        For lParseResultIdx = LBound(saEngParseResult) To UBound(saEngParseResult)
'            vOutRngArr(lRow, 2) = vOutRngArr(lRow, 2) + saEngParseResult(lParseResultIdx) + IIf(sSuffix = "", "", "_" + sSuffix) + _
'                                  IIf(lParseResultIdx = UBound(saEngParseResult), "", vbLf)
            vOutRngArr(lRow, 2) = vOutRngArr(lRow, 2) + saEngParseResult(lParseResultIdx) + sSuffix + _
                                  IIf(lParseResultIdx = UBound(saEngParseResult), "", vbLf)
        Next lParseResultIdx

        vOutRngArr(lRow, 3) = sGenType '속성명 점검결과
        vOutRngArr(lRow, 4) = IIf(sStdTermDataTypeSize <> "", sStdTermDataTypeSize, "") '표준용어 Type/Size
        sLastWord = Replace(Replace(SplitAndGetNItem(saKorParseResult(1), "_", -1), "[", ""), "]", "")
        vOutRngArr(lRow, 5) = sLastWord '속성명 종결어
        If Not oStdWordDic.Exists(sLastWord) Then
            sLastWordChk = "단어 없음"
            Set oStdWord = Nothing
        Else
            Set oStdWordObj = oStdWordDic.Item(sLastWord)
            If TypeOf oStdWordObj Is CStdWord Then
                Set oStdWord = oStdWordObj
            ElseIf TypeOf oStdWordObj Is CStdWordCol Then
                '논리명이 동일한 2개 이상의 단어가 존재할 때: 속성분류어로 지정된 단어개체를 찾아 보고 없으면 첫번째 단어개체로 지정
                For lIdx = 1 To oStdWordObj.Count
                    Set oStdWord = oStdWordObj.Items(lIdx)
                    If oStdWord.m_b속성분류어여부 = True Then Exit For
                Next lIdx
                If oStdWord.m_b속성분류어여부 = False Then Set oStdWord = oStdWordObj.Items(1)
            End If

            sLastWordChk = oStdWord.GetLastWordChk
        End If
        vOutRngArr(lRow, 6) = sLastWordChk '속성명 종결어 점검결과
        'vOutRngArr(lRow, 7) = oStdDomainDic.GetCheckAttrDataType(sLastWord, sAttrDataTypeSize, sStdTermDataTypeSize) '도메인, Data Type 점검결과
        vOutRngArr(lRow, 7) = oStdDomainDic.GetCheckAttrDataType(sLastWord, oStdWord, sAttrDataTypeSize, sStdTermDataTypeSize) '도메인, Data Type 점검결과
        vOutRngArr(lRow, 8) = Get추가후보단어(saKorParseResult(1)) '추가 후보 단어
SkipBlank:
    Next lRow

Finalize:
    oOutRange.Value2 = vOutRngArr
    Set oStdWordDic = Nothing
    Set oStdTermDic = Nothing
    Set oStdDomainDic = Nothing
    Set oWordMatchCol = Nothing
    Application.ScreenUpdating = True
End Sub

Public Function Get추가후보단어(a논리명조합 As String) As String
    If InStr(1, a논리명조합, "[") = 0 Then Exit Function '

    Dim sa논리명() As String, i As Integer, sToken As String, sWord As String
    Dim sWordList As String: sWordList = ""
    Dim iPrevTokenLen As Integer, bIsConcat As Boolean, bIsInitConcat As Boolean
    sa논리명 = Split(Replace(a논리명조합, vbLf, "_"), "_")
    For i = 0 To UBound(sa논리명)
        sToken = sa논리명(i)
        If Left(sToken, 1) = "[" Then
            sWord = Mid(sToken, 2, Len(sToken) - 2)
            bIsConcat = (iPrevTokenLen <= 1) And (Len(sWord) = 1) And (Not bIsInitConcat)
            sWordList = sWordList + IIf((sWordList > "") And Not bIsConcat, vbLf, "") + sWord
            bIsInitConcat = False
            iPrevTokenLen = Len(sWord)
        Else
            bIsInitConcat = True
        End If
    Next i
    Get추가후보단어 = sWordList
End Function

Public Sub 후보단어추가(aAttrBaseRange As Range, a후보단어BaseRange As Range)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim oWordCol As Collection, oAttrRange As Range, o후보단어Range As Range
    Dim lOffset As Long, sAttrName As String, s후보단어 As String, sa후보단어() As String, i As Integer
    Set oWordCol = New Collection
    Set o후보단어Range = a후보단어BaseRange
    Set oAttrRange = aAttrBaseRange

On Error Resume Next '중복 단어 추가시 오류 Skip
    '단어 Collection 생성
    For lOffset = 0 To 60000
        sAttrName = oAttrRange.Offset(lOffset, 0)
        If sAttrName = "" Then Exit For
        s후보단어 = o후보단어Range.Offset(lOffset, 0)
        sa후보단어 = Split(s후보단어, vbLf)
        For i = 0 To UBound(sa후보단어)
            s후보단어 = sa후보단어(i)
            If Trim(s후보단어) <> "" Then oWordCol.Add s후보단어, s후보단어
        Next i
    Next lOffset
On Error GoTo 0

    '단어사전 Sheet에 추가
    Dim oSht As Worksheet, oNumRange As Range, oWordRange As Range, dMatchResult As Double
    Set oSht = Worksheets("표준단어사전")
    Set oNumRange = oSht.Range("A1").End(xlDown) '순번(No)의 값이 있는 마지막 행(서식 복사용)
    Set oWordRange = oSht.Range("B1").End(xlDown) '단어논리명 컬럼의 값이 있는 마지막 행

    'oSht.Select
    'Range(oNumRange, oNumRange.End(xlToRight)).Select '순번(No)의 값이 있는 마지막 행(서식 복사용) 선택

On Error Resume Next 'Match 함수에서 찾는 값이 없을 때 발생하는 오류 무시
    '후보단어가 기존 단어목록에 존재하는지 확인
    For i = oWordCol.Count To 1 Step -1
        s후보단어 = oWordCol.Item(i)
        'if Application.WorksheetFunction.CountIf(
        dMatchResult = 0
        dMatchResult = Application.WorksheetFunction.Match(s후보단어, oSht.Range("B:B"), 0)
        If dMatchResult > 0 Then oWordCol.Remove (i) '이미 해당 후보단어가 목록에 있는 경우 Collection에서 삭제
    Next i
On Error GoTo 0

    If oWordCol.Count = 0 Then
        MsgBox "추가할 후보단어가 없습니다."
        GoTo Exit_Sub
    End If

    lOffset = 0
    oSht.Select
    oNumRange.EntireRow.Select '순번(No)의 값이 있는 마지막 행(서식 복사용) 선택
    Selection.Copy
    oWordRange.Offset(1, 0).Select
    'oWordRange.Resize(oWordCol.Count, 1).EntireRow.Select '추가 후보단어 갯수만큼 행 선택
    Selection.Resize(oWordCol.Count, 1).EntireRow.Select
    Selection.Insert Shift:=xlDown
    oWordRange.Offset(1, -1).Select
    Selection.Resize(oWordCol.Count, 5).ClearContents
    Application.CutCopyMode = False

    For i = 1 To oWordCol.Count
        s후보단어 = oWordCol.Item(i)
        lOffset = lOffset + 1
        oWordRange.Offset(lOffset, -1) = "추가" 'No 초기화
        oWordRange.Offset(lOffset, 0) = s후보단어 '단어논리명 초기화
        oWordRange.Offset(lOffset, 1).Formula = _
            "=CONCATENATE(""("", " & oWordRange.Offset(lOffset, 0).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ", "")"")" '단어물리명(추가시 논리명에 괄호 붙여서 기본생성)
        'oWordRange.Offset(lOffset, 2) = "" '단어영문명 초기화
        'oWordRange.Offset(lOffset, 3) = "" '단어설명 초기화
Skip_후보단어:
    Next i

    'Application.CutCopyMode = False
    oWordRange.Offset(lOffset, 0).Select
    i = oWordCol.Count

Exit_Sub:
    Set oWordCol = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If i > 0 Then MsgBox CStr(i) + "개의 후보단어를 추가하였습니다."
End Sub

'Public Sub 표준사전백업테스트()
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.DisplayAlerts = False
'    '단어사전백업
'    If Not IsSheetExists("표준단어사전_Bak") Then
'        Worksheets("표준단어사전").Copy After:=Worksheets(Worksheets.Count)
'        Worksheets(Worksheets.Count).Name = "표준단어사전_Bak"
'    Else
'        Worksheets("표준단어사전_Bak").Activate
'        Worksheets("표준단어사전_Bak").Range("A1").Select
'        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).EntireColumn.Delete
'        'DoClearList Worksheets("표준단어사전_Bak").Range("A2"), True
'
'        Worksheets("표준단어사전").Activate
'        Worksheets("표준단어사전").Range("A1").Select
'        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).EntireColumn.Select
'        Selection.Copy
'        Worksheets("표준단어사전_Bak").Activate
'        Worksheets("표준단어사전_Bak").Range("A1").Select
'        ActiveSheet.Paste
'    End If
'    Worksheets("표준단어사전_Bak").Range("A1").Select
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.DisplayAlerts = True
'End Sub

'Sheet 백업
Public Sub DoBackupSheet(aSht As Worksheet)
    Dim sOrgSheetName As String, sBakSheetName As String, oCurrentRange As Range
    sOrgSheetName = aSht.Name
    sBakSheetName = sOrgSheetName + "_Bak"
    If IsSheetExists(sBakSheetName) Then
        Application.DisplayAlerts = False
        Worksheets(sBakSheetName).Delete
        Application.DisplayAlerts = True
    End If
    Worksheets(sOrgSheetName).Copy After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = sBakSheetName
    Worksheets(sBakSheetName).Range("A1").Select
End Sub

'Private Sub ClearTest()
'    DoClearList Sheets("표준단어사전").Range("A2"), True
'End Sub

'Public Sub TestRefreshDateTime(a표준사전새로고침Range As Range)
'    Dim sPrevRefreshDateTime As String, lStartIdx As Long, lEndIdx As Long, lNumChar As Long
'    sPrevRefreshDateTime = Trim(a표준사전새로고침Range.Value2)
'    If sPrevRefreshDateTime <> "" Then
'        lStartIdx = InStr(1, sPrevRefreshDateTime, ": ") + 2
'        lEndIdx = InStr(1, sPrevRefreshDateTime, vbLf)
'        If lEndIdx = 0 Then lEndIdx = 31
'        lNumChar = lEndIdx - lStartIdx
'        sPrevRefreshDateTime = Mid(sPrevRefreshDateTime, lStartIdx, lNumChar)
'    End If
'
'    Dim sNewRefreshDateTime As String
'    sNewRefreshDateTime = "표준사전 기준일시: " + Format(Now, "yyyy-mm-dd hh:nn:ss") + vbLf + _
'                          "백업사전 기준일시: " + sPrevRefreshDateTime
'    Debug.Print sNewRefreshDateTime
'End Sub

Public Sub 표준사전새로고침(a표준사전새로고침Range As Range)
    Dim sPrevRefreshDateTime As String, lStartIdx As Long, lEndIdx As Long, lNumChar As Long
    sPrevRefreshDateTime = Trim(a표준사전새로고침Range.Value2)
    If sPrevRefreshDateTime <> "" Then
        lStartIdx = InStr(1, sPrevRefreshDateTime, ": ") + 2
        lEndIdx = InStr(1, sPrevRefreshDateTime, vbLf)
        If lEndIdx = 0 Then lEndIdx = 31
        lNumChar = lEndIdx - lStartIdx
        sPrevRefreshDateTime = Mid(sPrevRefreshDateTime, lStartIdx, lNumChar)
    End If
    Application.StatusBar = "표준사전 새로고침 시작..."
    Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual: Application.DisplayAlerts = False

    Dim sConnectionString As String, s표준단어사전Query As String, s표준용어사전Query As String, s표준도메인사전Query As String

    sConnectionString = Sheets("Config").Range("ConnectionString").Value2
    s표준단어사전Query = Sheets("Config").Range("표준단어사전Query").Value2
    s표준용어사전Query = Sheets("Config").Range("표준용어사전Query").Value2
    s표준도메인사전Query = Sheets("Config").Range("표준도메인사전Query").Value2

    Dim oDBCon As CDBConnection
    Set oDBCon = New CDBConnection
    oDBCon.InitProperty "표준사전", sConnectionString

    Dim oSht As Worksheet
    '표준단어 사전 갱신
    Set oSht = Sheets("표준단어사전")
    DoBackupSheet oSht
    DoClearList oSht.Range("A2"), True
    oDBCon.PopulateQueryResult s표준단어사전Query, oSht

    '표준용어 사전 갱신
    Set oSht = Sheets("표준용어사전")
    DoBackupSheet oSht
    DoClearList oSht.Range("A2"), True
    oDBCon.PopulateQueryResult s표준용어사전Query, oSht

    '표준도메인 사전 갱신
    Set oSht = Sheets("표준도메인사전")
    DoBackupSheet oSht
    DoClearList oSht.Range("A2"), True
    oDBCon.PopulateQueryResult s표준도메인사전Query, oSht

    Set oDBCon = Nothing
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: Application.DisplayAlerts = True
    Application.StatusBar = ""
    Dim sNewRefreshDateTime As String
    sNewRefreshDateTime = "표준사전 기준일시: " + Format(Now, "yyyy-mm-dd hh:nn:ss") + vbLf + _
                          "백업사전 기준일시: " + sPrevRefreshDateTime
    a표준사전새로고침Range.Value2 = sNewRefreshDateTime
    MsgBox "표준사전 새로고침 완료", vbInformation
End Sub

Public Sub SetInitParseResult(ByRef saParseResult As Variant, Optional a동음이의어Cnt As Long)
    Dim lIdx As Long, sParseToken As String, lInitBound As Long, lMultIdx As Long
    lInitBound = UBound(saParseResult)
    ReDim Preserve saParseResult(1 To UBound(saParseResult) * a동음이의어Cnt)

    Dim lTgtIdx As Long, lOffset As Long
    lOffset = lInitBound
    For lIdx = LBound(saParseResult) To lInitBound
        sParseToken = saParseResult(lIdx)
        For lMultIdx = 2 To a동음이의어Cnt
            lTgtIdx = lInitBound + lIdx + (lOffset * (lMultIdx - 2))
            saParseResult(lTgtIdx) = sParseToken
        Next lMultIdx
    Next lIdx
End Sub

Public Sub TestParseResult()
    Dim lIdx As Long, sParseToken As String, lInitBound As Long, lMultIdx As Long, a동음이의어Cnt As Long
    lInitBound = 2: a동음이의어Cnt = 3

    'ClearImmediateWindow
    Dim lTgtIdx As Long, lOffset As Long
    lOffset = lInitBound
    For lIdx = 1 To lInitBound
    'For lIdx = lInitBound To lInitBound * (a동음이의어Cnt - 1)
    'For lIdx = lInitBound + 1 To lInitBound * (a동음이의어Cnt)
        sParseToken = "test"
        For lMultIdx = 2 To a동음이의어Cnt
            lTgtIdx = lInitBound + lIdx + (lOffset * (lMultIdx - 2))
            Debug.Print CStr(lIdx) + " : " + CStr(lMultIdx) + " : " + CStr(lTgtIdx)
        Next lMultIdx
    Next lIdx

    Debug.Print
End Sub

Public Sub ClearImmediateWindow()
    Application.SendKeys "^g ^a {DEL}", True
    Application.SendKeys "{F7}", True
End Sub

Public Sub TestPartialParse()
    Dim lIdx As Long, lParseResultOffset As Long, lParseResultIdx As Long, lWordObjCnt As Long, lTgtIdx As Long, s단어물리명 As String
    lParseResultOffset = 3
    lWordObjCnt = 2
    lTgtIdx = 1
                    For lIdx = 1 To lParseResultOffset
                        'sParsingToken = saEngParseResult(1)
                        For lParseResultIdx = 1 To lWordObjCnt
                            s단어물리명 = "물리명" + CStr(lParseResultIdx)
                            'Set oStdWord = oStdWordObj.Items(lParseResultIdx)
                            'lTgtIdx = lParseResultIdx + (lParseResultOffset * (lParseResultIdx - 1))
                            'lTgtIdx = lParseResultIdx + lIdx + (lParseResultOffset * (lIdx - 1)) - 1
                            'saKorParseResult(lParseResultIdx) = saKorParseResult(lParseResultIdx) + IIf(saKorParseResult(lParseResultIdx) = "", "", "_") + oStdWord.m_s단어논리명
                            'saEngParseResult(lParseResultIdx) = saEngParseResult(lParseResultIdx) + IIf(saEngParseResult(lParseResultIdx) = "", "", "_") + oStdWord.m_s단어물리명
                            Debug.Print CStr(lIdx) + " : " + CStr(lParseResultIdx) + " : " + CStr(lTgtIdx) + " : " + s단어물리명
                            lTgtIdx = lTgtIdx + 1
                        Next lParseResultIdx
                    Next lIdx
    Debug.Print


    lTgtIdx = 1
                    For lParseResultIdx = 1 To lWordObjCnt
                        s단어물리명 = "물리명" + CStr(lParseResultIdx)
                        For lIdx = 1 To lParseResultOffset
                            Debug.Print CStr(lIdx) + " : " + CStr(lParseResultIdx) + " : " + CStr(lTgtIdx) + " : " + s단어물리명
                            lTgtIdx = lTgtIdx + 1
                        Next lIdx
                    Next lParseResultIdx

End Sub


' 속성명에 부적합한 문자가 포함되어 있는지 점검
Public Function IsValidAttributeName(aAttrName As String) As Boolean
    IsValidAttributeName = True
    'Dim aInvalidChar(1 To 2) As Variant
    Dim aInvalidChar As Variant, i As Integer
    aInvalidChar = Array("[", vbLf)
    For i = LBound(aInvalidChar) To UBound(aInvalidChar)
        If InStr(1, aAttrName, aInvalidChar(i)) > 0 Then
            IsValidAttributeName = False
            Exit Function
        End If
    Next
End Function

