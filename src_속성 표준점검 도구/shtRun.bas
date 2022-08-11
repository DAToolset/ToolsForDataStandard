VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub chkOnlyForSelectedAttr_Click()
    Selection.Select
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show (vbModal)
End Sub

Private Sub cmdAppendWord_Click()
    후보단어추가 aAttrBaseRange:=Names("속성목록Base").RefersToRange.Offset(1, 0) _
               , a후보단어BaseRange:=Names("후보단어목록Base").RefersToRange.Offset(1, 0)
End Sub

Private Sub cmdClear_Click()
    If IsOkToGo("표준 점검결과를 초기화합니다." + vbLf + "계속 진행하시겠습니까?") = False Then Exit Sub
    ClearResult
End Sub

Private Sub cmdConfig_Click()
    shtConfig.Activate
End Sub

Private Sub cmdRefreshStdDic_Click()
    '표준사전 갱신
    If IsOkToGo("표준사전을 새로고침합니다." + vbLf + "계속 진행하시겠습니까?") = False Then Exit Sub
    표준사전새로고침 Range("표준사전기준일시")
End Sub

Private Sub cmdRun_Click()
    Dim sSelectedRunMsg As String, bIsOnlyForSelectedAttr As Boolean, oSelectedAttrRange As Range
    If (chkOnlyForSelectedAttr.Value = True) Or IsShiftKeyDown Then
        sSelectedRunMsg = "★점검 대상: 선택한 속성★"
        bIsOnlyForSelectedAttr = True
        '다른 열을 선택했더라도 속성명, DataType/Len 컬럼 Range 구하기
        Set oSelectedAttrRange = Range("속성목록Base").Offset(1, 0)
        Dim lRowOffset As Long
        lRowOffset = Selection.Row - oSelectedAttrRange.Row
        Set oSelectedAttrRange = oSelectedAttrRange.Offset(lRowOffset, 0).Resize(Selection.Rows.Count, 2)
    Else
        sSelectedRunMsg = "점검 대상: 전체 속성"
        bIsOnlyForSelectedAttr = False
    End If

    If chkRefreshStdDic.Value = True Then
        If IsOkToGo("표준사전을 새로고침하고 표준 점검을 실행합니다." + vbLf + _
                    "(기존 표준사전은 백업)" + vbLf + _
                    sSelectedRunMsg + vbLf + vbLf + _
                    "★각 sheet에 설정된 필터를 해제합니다★" + vbLf + vbLf + _
                    "시간이 몇 분정도 걸릴 수 있습니다." + vbLf + _
                    "계속 진행하시겠습니까?") = False Then
            End
            Exit Sub
        End If
        표준사전새로고침 Range("표준사전기준일시")
        'Range("표준사전기준일시").Value2 = "표준사전 기준일시: " + Format(Now, "yyyy-mm-dd hh:nn:ss")
    Else
        If IsOkToGo("표준 점검을 실행합니다.(표준사전 유지)" + vbLf + _
                    sSelectedRunMsg + vbLf + vbLf + _
                    "★각 sheet에 설정된 필터를 해제합니다★" + vbLf + vbLf + _
                    "시간이 몇 분정도 걸릴 수 있습니다." + vbLf + _
                    "계속 진행하시겠습니까?") = False Then
            End
            Exit Sub
        End If
    End If
    Dim eWordMatchDirection As WordMatchDirection
    If (chkLtoR = False) And (chkRtoL = False) Then
        Call MsgBox("단어조합방향을 선택하세요.", vbCritical Or vbDefaultButton1, "오류")
        Exit Sub
    End If
    If chkLtoR Then eWordMatchDirection = LtoR
    If chkRtoL Then eWordMatchDirection = RtoL
    If (chkLtoR) And (chkRtoL) Then eWordMatchDirection = Both

    ClearResult bIsOnlyForSelectedAttr '표준점검결과 내용 삭제
    ClearAllFilters '각 Sheet의 Filter 해제

    Dim eStdDicMatchOption As StdDicMatchOption
    If optWordAndTerm = True Then
        eStdDicMatchOption = WordAndTerm
    ElseIf optWordOnly = True Then
        eStdDicMatchOption = WordOnly
    ElseIf optTermOnly = True Then
        eStdDicMatchOption = TermOnly
    End If

    '표준점검 Range("B3"), Range("D3"), Range("E3"), eStdDicMatchOption
    표준점검 aAttrBaseRange:=Range("속성목록Base").Offset(1, 0) _
             , aLNameBaseRange:=Range("표준단어논리명조합Base").Offset(1, 0) _
             , aPNameBaseRange:=Range("표준단어물리명조합Base").Offset(1, 0) _
             , aStdWordBaseRange:=Sheets("표준단어사전").Range("B2") _
             , aStdTermBaseRange:=Sheets("표준용어사전").Range("B2") _
             , aStdDomainBaseRange:=Sheets("표준도메인사전").Range("B2") _
             , aStdDicMatchOption:=eStdDicMatchOption _
             , aWordMatchDirection:=eWordMatchDirection _
             , aIsAllowDupWordLogicalName:=chkAllowDupWordLogicalName.Value _
             , aIsAllowDupWordPhysicalName:=chkAllowDupWordPhysicalName.Value _
             , aIsOnlyForSelectedAttr:=bIsOnlyForSelectedAttr _
             , aSelectedAttrRange:=oSelectedAttrRange
End Sub

'표준점검결과 내용 삭제 (점검 실행전 초기화)
Public Sub ClearResult(Optional aIsOnlyForSelectedAttr As Boolean = False)
    Dim oCurRange As Range: Set oCurRange = Selection 'ActiveCell
    Dim oBaseRange As Range
    Dim lColOffset As Long: lColOffset = 7 'Clear할 컬럼의 Offset(= 갯수-1)
    Dim lRowOffset As Long
    Application.ScreenUpdating = False
    'oBaseRange.Select
    Set oBaseRange = Range("표준단어논리명조합Base").Offset(1, 0)
    'Range(oBaseRange, oBaseRange.End(xlDown).Offset(, lColOffset)).ClearContents
    If aIsOnlyForSelectedAttr = False Then '전체 속성 점겅결과 삭제
        Range(oBaseRange, oBaseRange.Offset(shtRun.UsedRange.Rows.Count, lColOffset)).ClearContents
    Else '선택한 속성 점검결과만 삭제
        lRowOffset = oCurRange.Row - oBaseRange.Row
        Set oBaseRange = oBaseRange.Offset(lRowOffset, 0)
        Range(oBaseRange, oBaseRange.Offset(oCurRange.Rows.Count - 1, lColOffset)).ClearContents
    End If

    oCurRange.Select
    Application.ScreenUpdating = True
End Sub

'표준점검 실행전 각 Sheet의 Filter 해제
Private Sub ClearAllFilters()
    On Error Resume Next 'Filter가 설정되어 있지 않은 경우 오류가 발생하므로 Skip
    Sheets("속성 표준점검").AutoFilter.ShowAllData
    Sheets("표준단어사전").AutoFilter.ShowAllData
    Sheets("표준용어사전").AutoFilter.ShowAllData
    Sheets("표준도메인사전").AutoFilter.ShowAllData
    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate()
    ResetControlSize ActiveSheet
End Sub

Public Sub PartialRunTest()
    Dim lRow As Long, oSelRange As Range
    Set oSelRange = Selection
    lRow = oSelRange.Rows.Count
    Debug.Print lRow
    Debug.Print oSelRange.Rows(2).Address
End Sub
