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

Dim oStdWordDic As CStdWordDic '�ܾ����
Dim oStdTermDic As CStdTermDic '������
Dim oStdDomainDic As CStdDomainDic '�����λ���

Dim eStdDicMatchOption As StdDicMatchOption 'ǥ�ػ��� ã�� �ɼ�
Dim eWordMatchDirection As WordMatchDirection '�ܾ����� ����

'ǥ������ ���� ���ν���
'* Parameter
'  - aAttrBaseRange              : �Է�-���˴�� �Ӽ� ����� ���� ��ġ
'  - aLNameBaseRange             : ���-ǥ�شܾ� ���� ������ ���� ��ġ
'  - aPNameBaseRange             : ���-ǥ�شܾ� ������ ������ ���� ��ġ
'  - aStdWordBaseRange           : ����-ǥ�شܾ������ ���� ��ġ
'  - aStdTermBaseRange           : ����-ǥ�ؿ������� ���� ��ġ
'  - aStdDomainBaseRange         : ����-ǥ�ص����λ����� ���� ��ġ
'  - aStdDicMatchOption          : �ɼ�-ǥ�ػ���ã�� �ɼ� ���� ��(1:�ܾ�&���, 2:�ܾ�, 3:���)
'  - aWordMatchDirection         : �ɼ�-�ܾ����չ��� �ɼ� ���� ��(1:��->��, 2:��->��, 3:���)
'  - aIsAllowDupWordLogicalName  : �ɼ�-ǥ�شܾ� ���� �ߺ�(�������Ǿ�) ��� ����
'  - aIsAllowDupWordPhysicalName : �ɼ�-ǥ�شܾ� ������ �ߺ�(�������Ǿ�) ��� ����
'  - aIsOnlyForSelectedAttr      : �ɼ�-������ �Ӽ��� �������� ����(True: ������ �Ӽ��� ����, False: ��ü �Ӽ� ����)
'  - aSelectedAttrRange          : �ɼ�-������ �Ӽ��� ����� ����
Public Sub ǥ������(aAttrBaseRange As Range, aLNameBaseRange As Range, aPNameBaseRange As Range, _
        aStdWordBaseRange As Range, aStdTermBaseRange As Range, aStdDomainBaseRange As Range, _
        aStdDicMatchOption As StdDicMatchOption, aWordMatchDirection As WordMatchDirection, _
        aIsAllowDupWordLogicalName As Boolean, aIsAllowDupWordPhysicalName As Boolean, _
        aIsOnlyForSelectedAttr As Boolean, _
        aSelectedAttrRange As Range)

    Application.ScreenUpdating = False
    eStdDicMatchOption = aStdDicMatchOption
    eWordMatchDirection = aWordMatchDirection
    'GetDicMap
    '�ܾ���� �ҷ����� ---------------------------------------------------------------------
    Set oStdWordDic = New CStdWordDic: oStdWordDic.m_eStdDicMatchOption = aStdDicMatchOption
    oStdWordDic.Load aStdWordBaseRange, aIsAllowDupWordLogicalName, aIsAllowDupWordPhysicalName
    If oStdWordDic.Count = 0 Then
        MsgBox "�ܾ������ ��� �ֽ��ϴ�." + vbLf + "ǥ�� ������ �����մϴ�", vbCritical, "����"
        Exit Sub
    End If

    '������ �ҷ����� ---------------------------------------------------------------------
    Set oStdTermDic = New CStdTermDic: oStdTermDic.m_eStdDicMatchOption = aStdDicMatchOption
    oStdTermDic.SetStdWordDicP oStdWordDic.m_oStdWordDicP
    oStdTermDic.Load aStdTermBaseRange

    '�����λ��� �ҷ����� ---------------------------------------------------------------------
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
    Dim bWordMatched As Boolean, bTermMatched As Boolean, b�������Ǿ�Matched As Boolean, b�������Ǿ�Mached As Boolean
    Dim b��ǥ�شܾ�MatchedL As Boolean, b��ǥ�شܾ�MatchedR As Boolean
    Dim vInRngArr As Variant, vOutRngArr As Variant, lRow As Long, lLastRow As Long, oOutRange As Range
    Dim lParseResultOffset As Long, lTgtIdx As Long, lIdx As Long

    lRowOffset = 0
    '�Է¹��� �о Variant array�� ���
    If aAttrBaseRange.Offset(1, 0).Value2 = "" Then 'ù �ุ �����Ͱ� �ִ� ���
        vInRngArr = aAttrBaseRange.Resize(, 2).Value2 '�д� ����: 2�� �÷�
    Else
        If Not aIsOnlyForSelectedAttr Then '������ �Ӽ��� ���� �ʴ� ��� --> ��ü �б�
            vInRngArr = Range(aAttrBaseRange, aAttrBaseRange.End(xlDown)).Resize(, 2).Value2 '�д� ����: 2�� �÷�
        Else '������ �Ӽ��� �д� ���
            lRowOffset = aSelectedAttrRange.Row - aLNameBaseRange.Row
            vInRngArr = aSelectedAttrRange.Value2
        End If
    End If

    lLastRow = UBound(vInRngArr) 'aAttrBaseRange.End(xlDown).Row
    'Set oOutRange = aLNameBaseRange.Resize(lLastRow, 8) '���� ����: 8�� �÷�
    Set oOutRange = aLNameBaseRange.Offset(lRowOffset, 0).Resize(lLastRow, 8) '���� ����: 8�� �÷�
    'vOutRngArr = oOutRange.Value2
    ReDim vOutRngArr(1 To oOutRange.Rows.Count, 1 To oOutRange.Columns.Count)

    For lRow = LBound(vInRngArr) To UBound(vInRngArr)
        sAttrName = vInRngArr(lRow, 1)
        If Trim(sAttrName) = "" Then GoTo SkipBlank '������ �Ӽ����� ����ִ� ��� Skip
        sAttrDataTypeSize = vInRngArr(lRow, 2)
        bWordMatched = False: bTermMatched = False: lParseResultIdx = 0: b�������Ǿ�Matched = False
        b��ǥ�شܾ�MatchedL = False: b��ǥ�شܾ�MatchedR = False: b�������Ǿ�Mached = False

'        sKorParseResult = "": sEngParseResult = "": sStdTermDataTypeSize = ""
'        sKorParseResultR = "": sEngParseResultR = "": sSuffix = ""

        '------------------------------------------------------------------------------------------
        '���� �ʱ�ȭ
        ReDim saKorParseResult(1 To 1): ReDim saEngParseResult(1 To 1)
        ReDim saKorParseResultR(1 To 1): ReDim saEngParseResultR(1 To 1)
        sStdTermDataTypeSize = "": sSuffix = "": sGenType = "": sLastWord = "": sLastWordChk = ""
        '------------------------------------------------------------------------------------------

        If Not IsValidAttributeName(sAttrName) Then '�Ӽ��� �������� ���ڰ� ���Ե� ��� �޽��� �����ְ� Skip �Ǵ� �ߴ�
            If IsOkToGo("�Ӽ��� �������� ���ڰ� ���ԵǾ� �ֽ��ϴ�." + vbLf + _
                        "- �������� ����: [, ��и�����" + vbLf + vbLf + _
                        " - [Row#: " & CStr(lRow) + "] " & sAttrName + vbLf + vbLf + _
                        "�ش� �Ӽ��� �����ϰ� ��� �����Ͻðڽ��ϱ�?", "Ȯ��") Then
                saKorParseResult(1) = sAttrName
                sGenType = "�������� ���� Ȯ���ʿ�"
                GoTo Continue_OuterFor1
            Else
                MsgBox "������� ��û���� ������ �����մϴ�.", vbOKOnly + vbCritical, "����"
                GoTo Finalize
            End If
        End If

        '------------------------------------------------------------------------------------------
        '�Ӽ��� ���� ���� ����
        sSuffix = GetNumberSuffix(sAttrName)
        If sSuffix <> "" Then sAttrName = GetTextWithoutSuffix(sAttrName, sSuffix)
        '------------------------------------------------------------------------------------------

        '------------------------------------------------------------------------------------------
        '�Ӽ����� ���� �Ǵ� "_"�� ������ ��� ó��
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
                If oStdWordDic.Exists(sToken) Then 'Token�� ǥ�شܾ� ��Ͽ� ���Ե� ���
                    Set oStdWordObj = oStdWordDic.Item(sToken)
                    If TypeOf oStdWordObj Is CStdWord Then
                        '������ ������ ���
                        Set oStdWord = oStdWordObj
                        Set oStdWord = oStdWord.GetSWForNSW
                        If Not b��ǥ�شܾ�MatchedL Then
                            b��ǥ�شܾ�MatchedL = sToken <> oStdWord.m_s�ܾ����
                        End If
                        If oStdWord.m_b�������ߺ����� = True Then b�������Ǿ�Mached = True
                        saKorParseResult(1) = saKorParseResult(1) + IIf(saKorParseResult(1) = "", "", "_") + oStdWord.m_s�ܾ����
                        For lParseResultIdx = LBound(saEngParseResult) To UBound(saEngParseResult) '�������Ǿ��� n�� ������ ����
                            saEngParseResult(lParseResultIdx) = saEngParseResult(lParseResultIdx) + _
                                    IIf(saEngParseResult(lParseResultIdx) = "", "", "_") + oStdWord.m_s�ܾ����
                        Next lParseResultIdx
                    ElseIf TypeOf oStdWordObj Is CStdWordCol Then
                        '�������Ǿ �־� ������ 2�� �̻��� ���
                        b�������Ǿ�Matched = True
                        lParseResultOffset = UBound(saEngParseResult)
                        SetInitParseResult saEngParseResult, oStdWordObj.Count

                        Set oStdWord = oStdWordObj.Items(1)
                        saKorParseResult(1) = saKorParseResult(1) + IIf(saKorParseResult(1) = "", "", "_") + oStdWord.m_s�ܾ����

                        lTgtIdx = 1
                        For lParseResultIdx = 1 To oStdWordObj.Count
                            Set oStdWord = oStdWordObj.Items(lParseResultIdx)
                            For lIdx = 1 To lParseResultOffset
                                saEngParseResult(lTgtIdx) = saEngParseResult(lTgtIdx) + IIf(saEngParseResult(lTgtIdx) = "", "", "_") + oStdWord.m_s�ܾ����
                                lTgtIdx = lTgtIdx + 1
                            Next lIdx
                        Next lParseResultIdx
                    End If
                    bWordMatched = True
                Else '��Ī�� �ܾ ���� ���
                    sToken = IIf(saKorParseResult(1) = "", "[", "_[") + sToken + "]"
                    saKorParseResult(1) = saKorParseResult(1) + sToken
                    For lParseResultIdx = LBound(saEngParseResult) To UBound(saEngParseResult)
                        saEngParseResult(lParseResultIdx) = saEngParseResult(lParseResultIdx) + sToken
                    Next lParseResultIdx
                End If
Continue_InnerFor1:
            Next iFIdx
            sGenType = "ǥ�شܾ� " + IIf(bWordMatched, "����", "����") + "(����� ����)" _
                     + IIf(b�������Ǿ�Matched, vbLf + "(�������Ǿ� Ȯ���ʿ�)", "") _
                     + IIf(b��ǥ�شܾ�MatchedL, vbLf + "(��ǥ�شܾ� Ȯ���ʿ�)", "")
            GoTo SkipIfTermMatched
        End If
        '------------------------------------------------------------------------------------------

        iLenAttrName = Len(sAttrName)

        '------------------------------------------------------------------------------------------
        'ǥ�ؿ�� ã��(�ܾ�θ� ã�� ��찡 �ƴ� ��)
        If (Not eStdDicMatchOption = WordOnly) And _
           (oStdTermDic.Exists(sAttrName)) Then  '�Ӽ���� ��ġ�ϴ� ǥ�ؿ�� ����
            Set oStdTerm = oStdTermDic.Item(sAttrName)
            saKorParseResult(1) = oStdTerm.m_s�ܾ��������
            saEngParseResult(1) = oStdTerm.m_s������
            sStdTermDataTypeSize = oStdTerm.m_s������Ÿ�Ա��̸�
            sGenType = "ǥ�ؿ�� ��ġ"
            bTermMatched = True
            GoTo SkipIfTermMatched
        End If
        '------------------------------------------------------------------------------------------

        If eStdDicMatchOption = TermOnly Then
            sGenType = "ǥ�ؿ�� ����"
            GoTo SkipIfTermMatched
        End If

        If eWordMatchDirection = RtoL Then GoTo Skip_LtoR
        '------------------------------------------------------------------------------------------
        'ǥ�شܾ� �����ϱ� (�� --> �� Ž��)
        iFIdx = 1
        Do
            If iFIdx > iLenAttrName Then Exit Do
            Set oWordMatchCol = Nothing
            Set oWordMatchCol = New Collection
            For iTokenLen = 1 To iLenAttrName - iFIdx + 1 'iMaxStdWordLen
                'Token ����
                sToken = Mid(sAttrName, iFIdx, iTokenLen)
                If oStdWordDic.Exists(sToken) Then 'Token�� ǥ�شܾ� ��Ͽ� ���Ե� ���
                    Set oStdWordObj = oStdWordDic.Item(sToken)
                    oWordMatchCol.Add oStdWordObj
                End If
            Next iTokenLen
            If oWordMatchCol.Count > 0 Then '��Ī�� �ܾ �ִ� ���
                Set oStdWordObj = oWordMatchCol(oWordMatchCol.Count) '���� �� ǥ�شܾ� ��Ī��� ����
                If TypeOf oStdWordObj Is CStdWord Then
                    '������ ������ ���
                    Set oStdWord = oStdWordObj
                    sToken = oStdWord.m_s�ܾ����
                    Set oStdWord = oStdWord.GetSWForNSW
                    If Not b��ǥ�شܾ�MatchedL Then
                        b��ǥ�شܾ�MatchedL = sToken <> oStdWord.m_s�ܾ����
                    End If
                    If oStdWord.m_b�������ߺ����� = True Then b�������Ǿ�Mached = True
                    saKorParseResult(1) = saKorParseResult(1) + IIf(saKorParseResult(1) = "", "", "_") + oStdWord.m_s�ܾ����
                    For lParseResultIdx = LBound(saEngParseResult) To UBound(saEngParseResult) '�������Ǿ��� n�� ������ ����
                        saEngParseResult(lParseResultIdx) = saEngParseResult(lParseResultIdx) + _
                                IIf(saEngParseResult(lParseResultIdx) = "", "", "_") + oStdWord.m_s�ܾ����
                    Next lParseResultIdx

                ElseIf TypeOf oStdWordObj Is CStdWordCol Then
                    '�������Ǿ �־� ������ 2�� �̻��� ���
                    b�������Ǿ�Matched = True
                    lParseResultOffset = UBound(saEngParseResult)
                    SetInitParseResult saEngParseResult, oStdWordObj.Count

                    Set oStdWord = oStdWordObj.Items(1)
                    sToken = oStdWord.m_s�ܾ����
                    Set oStdWord = oStdWord.GetSWForNSW
                    If Not b��ǥ�شܾ�MatchedL Then
                        b��ǥ�شܾ�MatchedL = sToken <> oStdWord.m_s�ܾ����
                    End If
                    saKorParseResult(1) = saKorParseResult(1) + IIf(saKorParseResult(1) = "", "", "_") + oStdWord.m_s�ܾ����

                    lTgtIdx = 1
                    For lParseResultIdx = 1 To oStdWordObj.Count
                        Set oStdWord = oStdWordObj.Items(lParseResultIdx)
                        Set oStdWord = oStdWord.GetSWForNSW
                        For lIdx = 1 To lParseResultOffset
                            saEngParseResult(lTgtIdx) = saEngParseResult(lTgtIdx) + IIf(saEngParseResult(lTgtIdx) = "", "", "_") + oStdWord.m_s�ܾ����
                            lTgtIdx = lTgtIdx + 1
                        Next lIdx
                    Next lParseResultIdx
                End If

                'iFIdx = iFIdx + Len(oStdWord.m_s�ܾ����)
                iFIdx = iFIdx + Len(sToken)
                bWordMatched = True
            Else '��Ī�� �ܾ ���� ���
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
        'ǥ�شܾ� �����ϱ� (�� --> �� Ž��)
        iFIdx = iLenAttrName
        Do
            If iFIdx <= 0 Then Exit Do
            Set oWordMatchCol = Nothing
            Set oWordMatchCol = New Collection
            For iTokenLen = 1 To iFIdx
                'Token ����
                sToken = Mid(sAttrName, iFIdx - iTokenLen + 1, iTokenLen)
                If oStdWordDic.Exists(sToken) Then 'Token�� ǥ�شܾ� ��Ͽ� ���Ե� ���
                    Set oStdWordObj = oStdWordDic.Item(sToken)
                    oWordMatchCol.Add oStdWordObj
                End If
            Next iTokenLen
            If oWordMatchCol.Count > 0 Then '��Ī�� �ܾ �ִ� ���
                Set oStdWordObj = oWordMatchCol(oWordMatchCol.Count) '���� �� ǥ�شܾ� ��Ī��� ����
                If TypeOf oStdWordObj Is CStdWord Then
                    '������ ������ ���
                    Set oStdWord = oStdWordObj
                    sToken = oStdWord.m_s�ܾ����
                    Set oStdWord = oStdWord.GetSWForNSW
                    If Not b��ǥ�شܾ�MatchedR Then
                        b��ǥ�شܾ�MatchedR = sToken <> oStdWord.m_s�ܾ����
                    End If
                    If oStdWord.m_b�������ߺ����� = True Then b�������Ǿ�Mached = True
                    iFIdx = iFIdx - Len(sToken)
                    saKorParseResultR(1) = IIf(iFIdx > 0, "_", "") + oStdWord.m_s�ܾ���� + saKorParseResultR(1)
                    For lParseResultIdx = LBound(saEngParseResultR) To UBound(saEngParseResultR) '�������Ǿ��� n�� ������ ����
                        saEngParseResultR(lParseResultIdx) = IIf(iFIdx > 0, "_", "") + oStdWord.m_s�ܾ���� + saEngParseResultR(lParseResultIdx)
                    Next lParseResultIdx
                ElseIf TypeOf oStdWordObj Is CStdWordCol Then
                    '�������Ǿ �־� ������ 2�� �̻��� ���
                    b�������Ǿ�Matched = True
                    lParseResultOffset = UBound(saEngParseResultR)
                    SetInitParseResult saEngParseResultR, oStdWordObj.Count

                    Set oStdWord = oStdWordObj.Items(1)
                    sToken = oStdWord.m_s�ܾ����
                    Set oStdWord = oStdWord.GetSWForNSW
                    If Not b��ǥ�شܾ�MatchedR Then
                        b��ǥ�شܾ�MatchedR = sToken <> oStdWord.m_s�ܾ����
                    End If
                    iFIdx = iFIdx - Len(sToken)
                    saKorParseResultR(1) = IIf(iFIdx > 0, "_", "") + oStdWord.m_s�ܾ���� + saKorParseResultR(1)

                    lTgtIdx = 1
                    For lParseResultIdx = 1 To oStdWordObj.Count
                        Set oStdWord = oStdWordObj.Items(lParseResultIdx)
                        Set oStdWord = oStdWord.GetSWForNSW
                        For lIdx = 1 To lParseResultOffset
                            saEngParseResultR(lTgtIdx) = IIf(iFIdx > 0, "_", "") + oStdWord.m_s�ܾ���� + saEngParseResultR(lTgtIdx)
                            lTgtIdx = lTgtIdx + 1
                        Next lIdx
                    Next lParseResultIdx
                End If
                
                bWordMatched = True
            Else '��Ī�� �ܾ ���� ���
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
        sGenType = "ǥ�شܾ� " + IIf(bWordMatched, "����", "����") _
                 + IIf(b�������Ǿ�Matched, vbLf + "(�������Ǿ� Ȯ���ʿ�)", "") _
                 + IIf(b��ǥ�شܾ�MatchedL Or b��ǥ�شܾ�MatchedR, vbLf + "(��ǥ�شܾ� Ȯ���ʿ�)", "") _
                 + IIf(b�������Ǿ�Mached, vbLf + "(�������Ǿ� Ȯ���ʿ�)", "")
        If eWordMatchDirection = RtoL Then
            saKorParseResult(1) = saKorParseResultR(1)
            saEngParseResult(1) = saEngParseResultR(1)
        End If

        '��->�� ��Ī�� ��->�� ��Ī�� ����� �ٸ� ��� ��� ��������
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

            sGenType = sGenType + vbLf + "(���� ���� Ȯ���ʿ�)"
        End If

SkipIfTermMatched:
Continue_OuterFor1:
        'ǥ�شܾ� ���� ����
        For lParseResultIdx = LBound(saKorParseResult) To UBound(saKorParseResult)
'            vOutRngArr(lRow, 1) = vOutRngArr(lRow, 1) + saKorParseResult(lParseResultIdx) + IIf(sSuffix = "", "", "_" + sSuffix) + _
'                                  IIf(lParseResultIdx = UBound(saKorParseResult), "", vbLf)
            vOutRngArr(lRow, 1) = vOutRngArr(lRow, 1) + saKorParseResult(lParseResultIdx) + sSuffix + _
                                  IIf(lParseResultIdx = UBound(saKorParseResult), "", vbLf)
        Next lParseResultIdx

        'ǥ�شܾ� ������ ����
        For lParseResultIdx = LBound(saEngParseResult) To UBound(saEngParseResult)
'            vOutRngArr(lRow, 2) = vOutRngArr(lRow, 2) + saEngParseResult(lParseResultIdx) + IIf(sSuffix = "", "", "_" + sSuffix) + _
'                                  IIf(lParseResultIdx = UBound(saEngParseResult), "", vbLf)
            vOutRngArr(lRow, 2) = vOutRngArr(lRow, 2) + saEngParseResult(lParseResultIdx) + sSuffix + _
                                  IIf(lParseResultIdx = UBound(saEngParseResult), "", vbLf)
        Next lParseResultIdx

        vOutRngArr(lRow, 3) = sGenType '�Ӽ��� ���˰��
        vOutRngArr(lRow, 4) = IIf(sStdTermDataTypeSize <> "", sStdTermDataTypeSize, "") 'ǥ�ؿ�� Type/Size
        sLastWord = Replace(Replace(SplitAndGetNItem(saKorParseResult(1), "_", -1), "[", ""), "]", "")
        vOutRngArr(lRow, 5) = sLastWord '�Ӽ��� �����
        If Not oStdWordDic.Exists(sLastWord) Then
            sLastWordChk = "�ܾ� ����"
            Set oStdWord = Nothing
        Else
            Set oStdWordObj = oStdWordDic.Item(sLastWord)
            If TypeOf oStdWordObj Is CStdWord Then
                Set oStdWord = oStdWordObj
            ElseIf TypeOf oStdWordObj Is CStdWordCol Then
                '������ ������ 2�� �̻��� �ܾ ������ ��: �Ӽ��з���� ������ �ܾü�� ã�� ���� ������ ù��° �ܾü�� ����
                For lIdx = 1 To oStdWordObj.Count
                    Set oStdWord = oStdWordObj.Items(lIdx)
                    If oStdWord.m_b�Ӽ��з���� = True Then Exit For
                Next lIdx
                If oStdWord.m_b�Ӽ��з���� = False Then Set oStdWord = oStdWordObj.Items(1)
            End If

            sLastWordChk = oStdWord.GetLastWordChk
        End If
        vOutRngArr(lRow, 6) = sLastWordChk '�Ӽ��� ����� ���˰��
        'vOutRngArr(lRow, 7) = oStdDomainDic.GetCheckAttrDataType(sLastWord, sAttrDataTypeSize, sStdTermDataTypeSize) '������, Data Type ���˰��
        vOutRngArr(lRow, 7) = oStdDomainDic.GetCheckAttrDataType(sLastWord, oStdWord, sAttrDataTypeSize, sStdTermDataTypeSize) '������, Data Type ���˰��
        vOutRngArr(lRow, 8) = Get�߰��ĺ��ܾ�(saKorParseResult(1)) '�߰� �ĺ� �ܾ�
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

Public Function Get�߰��ĺ��ܾ�(a�������� As String) As String
    If InStr(1, a��������, "[") = 0 Then Exit Function '

    Dim sa����() As String, i As Integer, sToken As String, sWord As String
    Dim sWordList As String: sWordList = ""
    Dim iPrevTokenLen As Integer, bIsConcat As Boolean, bIsInitConcat As Boolean
    sa���� = Split(Replace(a��������, vbLf, "_"), "_")
    For i = 0 To UBound(sa����)
        sToken = sa����(i)
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
    Get�߰��ĺ��ܾ� = sWordList
End Function

Public Sub �ĺ��ܾ��߰�(aAttrBaseRange As Range, a�ĺ��ܾ�BaseRange As Range)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim oWordCol As Collection, oAttrRange As Range, o�ĺ��ܾ�Range As Range
    Dim lOffset As Long, sAttrName As String, s�ĺ��ܾ� As String, sa�ĺ��ܾ�() As String, i As Integer
    Set oWordCol = New Collection
    Set o�ĺ��ܾ�Range = a�ĺ��ܾ�BaseRange
    Set oAttrRange = aAttrBaseRange

On Error Resume Next '�ߺ� �ܾ� �߰��� ���� Skip
    '�ܾ� Collection ����
    For lOffset = 0 To 60000
        sAttrName = oAttrRange.Offset(lOffset, 0)
        If sAttrName = "" Then Exit For
        s�ĺ��ܾ� = o�ĺ��ܾ�Range.Offset(lOffset, 0)
        sa�ĺ��ܾ� = Split(s�ĺ��ܾ�, vbLf)
        For i = 0 To UBound(sa�ĺ��ܾ�)
            s�ĺ��ܾ� = sa�ĺ��ܾ�(i)
            If Trim(s�ĺ��ܾ�) <> "" Then oWordCol.Add s�ĺ��ܾ�, s�ĺ��ܾ�
        Next i
    Next lOffset
On Error GoTo 0

    '�ܾ���� Sheet�� �߰�
    Dim oSht As Worksheet, oNumRange As Range, oWordRange As Range, dMatchResult As Double
    Set oSht = Worksheets("ǥ�شܾ����")
    Set oNumRange = oSht.Range("A1").End(xlDown) '����(No)�� ���� �ִ� ������ ��(���� �����)
    Set oWordRange = oSht.Range("B1").End(xlDown) '�ܾ���� �÷��� ���� �ִ� ������ ��

    'oSht.Select
    'Range(oNumRange, oNumRange.End(xlToRight)).Select '����(No)�� ���� �ִ� ������ ��(���� �����) ����

On Error Resume Next 'Match �Լ����� ã�� ���� ���� �� �߻��ϴ� ���� ����
    '�ĺ��ܾ ���� �ܾ��Ͽ� �����ϴ��� Ȯ��
    For i = oWordCol.Count To 1 Step -1
        s�ĺ��ܾ� = oWordCol.Item(i)
        'if Application.WorksheetFunction.CountIf(
        dMatchResult = 0
        dMatchResult = Application.WorksheetFunction.Match(s�ĺ��ܾ�, oSht.Range("B:B"), 0)
        If dMatchResult > 0 Then oWordCol.Remove (i) '�̹� �ش� �ĺ��ܾ ��Ͽ� �ִ� ��� Collection���� ����
    Next i
On Error GoTo 0

    If oWordCol.Count = 0 Then
        MsgBox "�߰��� �ĺ��ܾ �����ϴ�."
        GoTo Exit_Sub
    End If

    lOffset = 0
    oSht.Select
    oNumRange.EntireRow.Select '����(No)�� ���� �ִ� ������ ��(���� �����) ����
    Selection.Copy
    oWordRange.Offset(1, 0).Select
    'oWordRange.Resize(oWordCol.Count, 1).EntireRow.Select '�߰� �ĺ��ܾ� ������ŭ �� ����
    Selection.Resize(oWordCol.Count, 1).EntireRow.Select
    Selection.Insert Shift:=xlDown
    oWordRange.Offset(1, -1).Select
    Selection.Resize(oWordCol.Count, 5).ClearContents
    Application.CutCopyMode = False

    For i = 1 To oWordCol.Count
        s�ĺ��ܾ� = oWordCol.Item(i)
        lOffset = lOffset + 1
        oWordRange.Offset(lOffset, -1) = "�߰�" 'No �ʱ�ȭ
        oWordRange.Offset(lOffset, 0) = s�ĺ��ܾ� '�ܾ���� �ʱ�ȭ
        oWordRange.Offset(lOffset, 1).Formula = _
            "=CONCATENATE(""("", " & oWordRange.Offset(lOffset, 0).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ", "")"")" '�ܾ����(�߰��� ���� ��ȣ �ٿ��� �⺻����)
        'oWordRange.Offset(lOffset, 2) = "" '�ܾ���� �ʱ�ȭ
        'oWordRange.Offset(lOffset, 3) = "" '�ܾ�� �ʱ�ȭ
Skip_�ĺ��ܾ�:
    Next i

    'Application.CutCopyMode = False
    oWordRange.Offset(lOffset, 0).Select
    i = oWordCol.Count

Exit_Sub:
    Set oWordCol = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If i > 0 Then MsgBox CStr(i) + "���� �ĺ��ܾ �߰��Ͽ����ϴ�."
End Sub

'Public Sub ǥ�ػ�������׽�Ʈ()
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.DisplayAlerts = False
'    '�ܾ�������
'    If Not IsSheetExists("ǥ�شܾ����_Bak") Then
'        Worksheets("ǥ�شܾ����").Copy After:=Worksheets(Worksheets.Count)
'        Worksheets(Worksheets.Count).Name = "ǥ�شܾ����_Bak"
'    Else
'        Worksheets("ǥ�شܾ����_Bak").Activate
'        Worksheets("ǥ�شܾ����_Bak").Range("A1").Select
'        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).EntireColumn.Delete
'        'DoClearList Worksheets("ǥ�شܾ����_Bak").Range("A2"), True
'
'        Worksheets("ǥ�شܾ����").Activate
'        Worksheets("ǥ�شܾ����").Range("A1").Select
'        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).EntireColumn.Select
'        Selection.Copy
'        Worksheets("ǥ�شܾ����_Bak").Activate
'        Worksheets("ǥ�شܾ����_Bak").Range("A1").Select
'        ActiveSheet.Paste
'    End If
'    Worksheets("ǥ�شܾ����_Bak").Range("A1").Select
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.DisplayAlerts = True
'End Sub

'Sheet ���
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
'    DoClearList Sheets("ǥ�شܾ����").Range("A2"), True
'End Sub

'Public Sub TestRefreshDateTime(aǥ�ػ������ΰ�ħRange As Range)
'    Dim sPrevRefreshDateTime As String, lStartIdx As Long, lEndIdx As Long, lNumChar As Long
'    sPrevRefreshDateTime = Trim(aǥ�ػ������ΰ�ħRange.Value2)
'    If sPrevRefreshDateTime <> "" Then
'        lStartIdx = InStr(1, sPrevRefreshDateTime, ": ") + 2
'        lEndIdx = InStr(1, sPrevRefreshDateTime, vbLf)
'        If lEndIdx = 0 Then lEndIdx = 31
'        lNumChar = lEndIdx - lStartIdx
'        sPrevRefreshDateTime = Mid(sPrevRefreshDateTime, lStartIdx, lNumChar)
'    End If
'
'    Dim sNewRefreshDateTime As String
'    sNewRefreshDateTime = "ǥ�ػ��� �����Ͻ�: " + Format(Now, "yyyy-mm-dd hh:nn:ss") + vbLf + _
'                          "������� �����Ͻ�: " + sPrevRefreshDateTime
'    Debug.Print sNewRefreshDateTime
'End Sub

Public Sub ǥ�ػ������ΰ�ħ(aǥ�ػ������ΰ�ħRange As Range)
    Dim sPrevRefreshDateTime As String, lStartIdx As Long, lEndIdx As Long, lNumChar As Long
    sPrevRefreshDateTime = Trim(aǥ�ػ������ΰ�ħRange.Value2)
    If sPrevRefreshDateTime <> "" Then
        lStartIdx = InStr(1, sPrevRefreshDateTime, ": ") + 2
        lEndIdx = InStr(1, sPrevRefreshDateTime, vbLf)
        If lEndIdx = 0 Then lEndIdx = 31
        lNumChar = lEndIdx - lStartIdx
        sPrevRefreshDateTime = Mid(sPrevRefreshDateTime, lStartIdx, lNumChar)
    End If
    Application.StatusBar = "ǥ�ػ��� ���ΰ�ħ ����..."
    Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual: Application.DisplayAlerts = False

    Dim sConnectionString As String, sǥ�شܾ����Query As String, sǥ�ؿ�����Query As String, sǥ�ص����λ���Query As String

    sConnectionString = Sheets("Config").Range("ConnectionString").Value2
    sǥ�شܾ����Query = Sheets("Config").Range("ǥ�شܾ����Query").Value2
    sǥ�ؿ�����Query = Sheets("Config").Range("ǥ�ؿ�����Query").Value2
    sǥ�ص����λ���Query = Sheets("Config").Range("ǥ�ص����λ���Query").Value2

    Dim oDBCon As CDBConnection
    Set oDBCon = New CDBConnection
    oDBCon.InitProperty "ǥ�ػ���", sConnectionString

    Dim oSht As Worksheet
    'ǥ�شܾ� ���� ����
    Set oSht = Sheets("ǥ�شܾ����")
    DoBackupSheet oSht
    DoClearList oSht.Range("A2"), True
    oDBCon.PopulateQueryResult sǥ�شܾ����Query, oSht

    'ǥ�ؿ�� ���� ����
    Set oSht = Sheets("ǥ�ؿ�����")
    DoBackupSheet oSht
    DoClearList oSht.Range("A2"), True
    oDBCon.PopulateQueryResult sǥ�ؿ�����Query, oSht

    'ǥ�ص����� ���� ����
    Set oSht = Sheets("ǥ�ص����λ���")
    DoBackupSheet oSht
    DoClearList oSht.Range("A2"), True
    oDBCon.PopulateQueryResult sǥ�ص����λ���Query, oSht

    Set oDBCon = Nothing
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: Application.DisplayAlerts = True
    Application.StatusBar = ""
    Dim sNewRefreshDateTime As String
    sNewRefreshDateTime = "ǥ�ػ��� �����Ͻ�: " + Format(Now, "yyyy-mm-dd hh:nn:ss") + vbLf + _
                          "������� �����Ͻ�: " + sPrevRefreshDateTime
    aǥ�ػ������ΰ�ħRange.Value2 = sNewRefreshDateTime
    MsgBox "ǥ�ػ��� ���ΰ�ħ �Ϸ�", vbInformation
End Sub

Public Sub SetInitParseResult(ByRef saParseResult As Variant, Optional a�������Ǿ�Cnt As Long)
    Dim lIdx As Long, sParseToken As String, lInitBound As Long, lMultIdx As Long
    lInitBound = UBound(saParseResult)
    ReDim Preserve saParseResult(1 To UBound(saParseResult) * a�������Ǿ�Cnt)

    Dim lTgtIdx As Long, lOffset As Long
    lOffset = lInitBound
    For lIdx = LBound(saParseResult) To lInitBound
        sParseToken = saParseResult(lIdx)
        For lMultIdx = 2 To a�������Ǿ�Cnt
            lTgtIdx = lInitBound + lIdx + (lOffset * (lMultIdx - 2))
            saParseResult(lTgtIdx) = sParseToken
        Next lMultIdx
    Next lIdx
End Sub

Public Sub TestParseResult()
    Dim lIdx As Long, sParseToken As String, lInitBound As Long, lMultIdx As Long, a�������Ǿ�Cnt As Long
    lInitBound = 2: a�������Ǿ�Cnt = 3

    'ClearImmediateWindow
    Dim lTgtIdx As Long, lOffset As Long
    lOffset = lInitBound
    For lIdx = 1 To lInitBound
    'For lIdx = lInitBound To lInitBound * (a�������Ǿ�Cnt - 1)
    'For lIdx = lInitBound + 1 To lInitBound * (a�������Ǿ�Cnt)
        sParseToken = "test"
        For lMultIdx = 2 To a�������Ǿ�Cnt
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
    Dim lIdx As Long, lParseResultOffset As Long, lParseResultIdx As Long, lWordObjCnt As Long, lTgtIdx As Long, s�ܾ���� As String
    lParseResultOffset = 3
    lWordObjCnt = 2
    lTgtIdx = 1
                    For lIdx = 1 To lParseResultOffset
                        'sParsingToken = saEngParseResult(1)
                        For lParseResultIdx = 1 To lWordObjCnt
                            s�ܾ���� = "������" + CStr(lParseResultIdx)
                            'Set oStdWord = oStdWordObj.Items(lParseResultIdx)
                            'lTgtIdx = lParseResultIdx + (lParseResultOffset * (lParseResultIdx - 1))
                            'lTgtIdx = lParseResultIdx + lIdx + (lParseResultOffset * (lIdx - 1)) - 1
                            'saKorParseResult(lParseResultIdx) = saKorParseResult(lParseResultIdx) + IIf(saKorParseResult(lParseResultIdx) = "", "", "_") + oStdWord.m_s�ܾ����
                            'saEngParseResult(lParseResultIdx) = saEngParseResult(lParseResultIdx) + IIf(saEngParseResult(lParseResultIdx) = "", "", "_") + oStdWord.m_s�ܾ����
                            Debug.Print CStr(lIdx) + " : " + CStr(lParseResultIdx) + " : " + CStr(lTgtIdx) + " : " + s�ܾ����
                            lTgtIdx = lTgtIdx + 1
                        Next lParseResultIdx
                    Next lIdx
    Debug.Print


    lTgtIdx = 1
                    For lParseResultIdx = 1 To lWordObjCnt
                        s�ܾ���� = "������" + CStr(lParseResultIdx)
                        For lIdx = 1 To lParseResultOffset
                            Debug.Print CStr(lIdx) + " : " + CStr(lParseResultIdx) + " : " + CStr(lTgtIdx) + " : " + s�ܾ����
                            lTgtIdx = lTgtIdx + 1
                        Next lIdx
                    Next lParseResultIdx

End Sub


' �Ӽ��� �������� ���ڰ� ���ԵǾ� �ִ��� ����
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

