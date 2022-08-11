VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdTermDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'표준용어 Dictionary Class
Option Explicit

Public m_oStdTermDic As Dictionary 'Key: 용어논리명, Value: CStdTerm instance
Private m_oStdWordDicP As Dictionary '단어사전(Key:단어물리명, Value: CStdWord instance)
Public m_o논리명조합Dic As Dictionary '용어의 논리명조합 동일성 판별을 위한 dictionary(Key:단어논리명조합, Value: Collection of CStdTerm)
Public m_eStdDicMatchOption As StdDicMatchOption
Public m_s논리명중복점검결과 As String

Private Sub Class_Initialize()
    Set m_oStdTermDic = New Dictionary
    Set m_o논리명조합Dic = New Dictionary
End Sub

Private Sub Class_Terminate()
    m_oStdTermDic.RemoveAll: Set m_oStdTermDic = Nothing
    m_o논리명조합Dic.RemoveAll: Set m_o논리명조합Dic = Nothing
End Sub

Public Sub SetStdWordDicP(aStdWordDicP As Dictionary)
    Set m_oStdWordDicP = aStdWordDicP
End Sub

'용어사전 Loading(By Variant Array)
Public Sub Load(aBaseRange As Range)
    '단어만 조합하는 경우 용어사전 Loading하지 않음
    If m_eStdDicMatchOption = WordOnly Then Exit Sub

    '목록에 아무 값이 없는 경우 exit
    If Trim(aBaseRange.Offset(1, 0)) = "" Then Exit Sub

    Dim vRngArr As Variant
    vRngArr = Range(aBaseRange, aBaseRange.End(xlDown)).Resize(, 11).Value2 '읽는 범위: 10개 컬럼
    Dim lRow As Long, sInfoMsg As String

    Dim oStdTerm As CStdTerm, oStdTermTmp As CStdTerm, oStdTermCol As Collection
    For lRow = LBound(vRngArr) To UBound(vRngArr)
        Set oStdTerm = New CStdTerm
        oStdTerm.m_s용어논리명 = vRngArr(lRow, 1)
        oStdTerm.m_s단어논리명조합 = vRngArr(lRow, 2)
        oStdTerm.m_s용어물리명 = vRngArr(lRow, 3)
        oStdTerm.m_s용어설명 = vRngArr(lRow, 4)
        oStdTerm.m_s도메인논리명 = vRngArr(lRow, 5)
        oStdTerm.m_s데이터타입명 = vRngArr(lRow, 6)
        oStdTerm.m_i길이 = CInt(vRngArr(lRow, 7))
        oStdTerm.m_i정도 = CInt(vRngArr(lRow, 8))
        oStdTerm.m_s정의업무 = vRngArr(lRow, 9)
        oStdTerm.m_s데이터타입길이명 = vRngArr(lRow, 10)
'        oStdTerm.m_s용어유형 = vRngArr(lRow, 8)
'        oStdTerm.m_s담당업무 = vRngArr(lRow, 9)
'        oStdTerm.m_s표준단어조합 = vRngArr(lRow, 10)
'        oStdTerm.m_s데이터타입길이명 = vRngArr(lRow, 11)
        oStdTerm.SetData m_oStdWordDicP

On Error Resume Next
        m_oStdTermDic.Add oStdTerm.m_s용어논리명, oStdTerm
        If Err <> 0 Then '용어 논리명 중복 발생
            Set oStdTermTmp = m_oStdTermDic.Item(oStdTerm.m_s용어논리명)
            sInfoMsg = "다음의 이유로 실행이 중단되었습니다: " & "용어논리명 중복" & vbLf & _
                       "▶ 중복 항목" & vbLf & _
                           vbTab & "용어논리명: " & oStdTerm.m_s용어논리명 & vbLf & _
                           vbTab & "용어물리명: " & oStdTerm.m_s용어물리명 & vbLf & _
                       "▶ 기존 항목" & vbLf & _
                           vbTab & "용어논리명: " & oStdTermTmp.m_s용어논리명 & vbLf & _
                           vbTab & "용어물리명: " & oStdTermTmp.m_s용어물리명 & vbLf & vbLf & _
                       "중복을 무시하고 계속 진행하시겠습니까?"

            If IsOkToGo(sInfoMsg, "용어 논리명 중복") Then
                Resume Next
            Else
                End
                Exit For
            End If
        End If

        Err.Clear

        '2021-02-07 추가
        If m_o논리명조합Dic.Exists(oStdTerm.m_sSorted단어논리명조합) Then
            '용어의 m_sSorted단어논리명조합으로 Item이 존재하는 경우
            Set oStdTermCol = m_o논리명조합Dic.Item(oStdTerm.m_sSorted단어논리명조합)
        Else '존재하지 않는 경우
            Set oStdTermCol = New Collection
            m_o논리명조합Dic.Add oStdTerm.m_sSorted단어논리명조합, oStdTermCol
        End If
        oStdTermCol.Add oStdTerm
    Next lRow

    Dim v논리명조합 As Variant, sChkResultTemp As String, sChkResult As String
    For Each v논리명조합 In m_o논리명조합Dic.Keys()
        Set oStdTermCol = m_o논리명조합Dic.Item(v논리명조합)
        sChkResultTemp = ""
        If oStdTermCol.Count > 1 Then
            lRow = 0
            Set oStdTerm = oStdTermCol.Item(1)
            sChkResultTemp = "용어논리명: " + oStdTerm.m_s용어논리명
            For Each oStdTerm In oStdTermCol
                lRow = lRow + 1
                sChkResultTemp = sChkResultTemp + vbLf + _
                                 "  - 용어물리명(" + CStr(lRow) + "): " + oStdTerm.m_s용어물리명
            Next oStdTerm
            sChkResult = sChkResult + sChkResultTemp + vbLf
        End If
    Next v논리명조합
    m_s논리명중복점검결과 = sChkResult
End Sub

Public Function Exists(aTerm As String) As Boolean
    Exists = m_oStdTermDic.Exists(aTerm)
End Function

Public Function Item(aTerm As String) As CStdTerm
    Set Item = m_oStdTermDic.Item(aTerm)
End Function

