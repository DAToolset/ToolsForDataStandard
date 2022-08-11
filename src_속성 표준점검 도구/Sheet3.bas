VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdRun_Click()
    Dim aAttrBaseRange As Range: Set aAttrBaseRange = Range("C4")
    Dim aStdWordBaseRange As Range: Set aStdWordBaseRange = Sheets("표준단어사전").Range("B2")
    Dim aIsAllowDupWordLogicalName As Boolean: aIsAllowDupWordLogicalName = True
    Dim aIsAllowDupWordPhysicalName As Boolean: aIsAllowDupWordPhysicalName = True
    Dim aIsOnlyForSelectedAttr As Boolean: aIsOnlyForSelectedAttr = False
    Dim aSelectedAttrRange As Range

    Dim oStdWordDic As CStdWordDic '단어사전
    Dim oStdTermDic As CStdTermDic '용어사전
    Dim oStdDomainDic As CStdDomainDic '도메인사전
    Dim vInRngArr As Variant, vOutRngArr As Variant, lRow As Long, lLastRow As Long, oOutRange As Range

'    Dim eStdDicMatchOption As StdDicMatchOption '표준사전 찾기 옵션
'    Dim eWordMatchDirection As WordMatchDirection '단어조합 방향

    '단어사전 불러오기 ---------------------------------------------------------------------
    Set oStdWordDic = New CStdWordDic: oStdWordDic.m_eStdDicMatchOption = WordAndTerm
    oStdWordDic.Load aStdWordBaseRange, aIsAllowDupWordLogicalName, aIsAllowDupWordPhysicalName
    If oStdWordDic.Count = 0 Then
        MsgBox "단어사전이 비어 있습니다." + vbLf + "표준 점검을 중지합니다", vbCritical, "오류"
        Exit Sub
    End If

    Dim oAttrRange As Range, lRowOffset As Long
    Dim sColName As String, sAttrName As String, iLenAttrName As Integer, sAttrDataTypeSize As String
    lRowOffset = 0
    '입력범위 읽어서 Variant array에 담기
    If aAttrBaseRange.Offset(1, 0).Value2 = "" Then '첫 행만 데이터가 있는 경우
        vInRngArr = aAttrBaseRange.Resize(, 2).Value2 '읽는 범위: 2개 컬럼
    Else
        If Not aIsOnlyForSelectedAttr Then '선택한 속성만 읽지 않는 경우 --> 전체 읽기
            vInRngArr = Range(aAttrBaseRange, aAttrBaseRange.End(xlDown)).Resize(, 2).Value2 '읽는 범위: 2개 컬럼
        Else '선택한 속성만 읽는 경우
'            lRowOffset = aSelectedAttrRange.Row - aLNameBaseRange.Row
'            vInRngArr = aSelectedAttrRange.Value2
        End If
    End If
    Set oOutRange = Range("F4")

    '------------------------------------------------------------------------------------------
    '물리명 기준으로 논리명 찾기
    Dim aColNamePart() As String, lColNamePartIdx As Long, oColNamePartAsStdWord As Collection
    Dim oStdWordObj As Object, oStdWord As CStdWord, sLName As String
    For lRow = LBound(vInRngArr) To UBound(vInRngArr)
        sColName = vInRngArr(lRow, 1)
        If Trim(sColName) = "" Then GoTo SkipBlank '점검할 속성명이 비어있는 경우 Skip
        aColNamePart = Split(sColName, "_")
        Set oColNamePartAsStdWord = Nothing
        Set oColNamePartAsStdWord = New Collection
        sLName = ""
        For lColNamePartIdx = LBound(aColNamePart) To UBound(aColNamePart)
            Set oStdWordObj = oStdWordDic.ItemP(aColNamePart(lColNamePartIdx))
            oColNamePartAsStdWord.Add oStdWordObj
            If TypeOf oStdWordObj Is CStdWord Then
                Set oStdWord = oStdWordObj
                sLName = sLName + IIf(sLName = "", "", "_") + oStdWord.m_s단어논리명
            End If
            oOutRange.Offset(lRow - 1, 0).Value2 = sLName
        Next lColNamePartIdx
SkipBlank:
    Next lRow
    '------------------------------------------------------------------------------------------
End Sub
