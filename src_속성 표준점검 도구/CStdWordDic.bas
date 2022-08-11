VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdWordDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'표준단어 Dictionary Class
Option Explicit

Public m_oStdWordDic As Dictionary 'Key: 단어논리명, Value: CStdWord Collection(2018-05-07) '변경전: CStdWord instance
Public m_oStdWordDicP As Dictionary 'Key: 단어물리명, Value: CStdWord Collection(2018-05-07) '변경전: CStdWord instance

Public m_eStdDicMatchOption As StdDicMatchOption

Private Sub Class_Initialize()
    Set m_oStdWordDic = New Dictionary
    Set m_oStdWordDicP = New Dictionary
End Sub

Private Sub Class_Terminate()
    m_oStdWordDic.RemoveAll
    m_oStdWordDicP.RemoveAll

    Set m_oStdWordDic = Nothing
    Set m_oStdWordDicP = Nothing
End Sub

'단어사전 Loading(By Variant Array)
Public Sub Load(aBaseRange As Range, aIsAllowDupWordLogicalName As Boolean, aIsAllowDupWordPhysicalName As Boolean)
On Error Resume Next
'    '용어만 찾는 경우 단어사전 Loading하지 않음
'    If m_eStdDicMatchOption = TermOnly Then Exit Sub

    '목록에 아무 값이 없는 경우 exit
    If Trim(aBaseRange.Offset(1, 0)) = "" Then Exit Sub

    Dim vRngArr As Variant
    vRngArr = Range(aBaseRange, aBaseRange.End(xlDown)).Resize(, 9).Value2 '읽는 범위: 6개 컬럼
    Dim lRow As Long, oStdWord As CStdWord, oStdWordTmp As CStdWord, sInfoMsg As String
    Dim bIs단어논리명존재 As Boolean, oStdWordCol As CStdWordCol, oExistedStdObj As Object

    For lRow = LBound(vRngArr) To UBound(vRngArr)
        bIs단어논리명존재 = False
        Set oStdWord = New CStdWord
        oStdWord.SetStdWordDic Me
        oStdWord.m_s단어논리명 = vRngArr(lRow, 1)
        oStdWord.m_s단어물리명 = vRngArr(lRow, 2)
        oStdWord.m_s단어영문명 = vRngArr(lRow, 3)
        oStdWord.m_s단어설명 = vRngArr(lRow, 4)
        'oStdWord.m_s단어유형 = vRngArr(lRow, 5)
        oStdWord.m_b표준여부 = GetBoolean(vRngArr(lRow, 5))
        oStdWord.m_b속성분류어여부 = GetBoolean(vRngArr(lRow, 6))
        oStdWord.m_s표준논리명 = vRngArr(lRow, 7)
        oStdWord.m_s동의어 = vRngArr(lRow, 8)
        oStdWord.m_s도메인분류명 = vRngArr(lRow, 9)

        '단어논리명 기준 Dictionary Build
        bIs단어논리명존재 = m_oStdWordDic.Exists(oStdWord.m_s단어논리명)
        If bIs단어논리명존재 Then
            If aIsAllowDupWordLogicalName Then
                '단어 논리명 중복(동음이의어) 허용
                If (oStdWord.m_b표준여부 = False) And _
                   (oStdWord.m_s단어논리명 <> oStdWord.m_s표준논리명) Then
                    GoTo Skip_단어논리명Build '표준여부 False이고 논리명과 표준논리명이 동일하지 않은 경우 Skip
                End If

                oStdWord.m_b논리명중복여부 = True
                '메시지 보여주지 않고 Collection만들어서 기존 Value item 대체
                Set oExistedStdObj = m_oStdWordDic.Item(oStdWord.m_s단어논리명) '기존 단어
                If TypeOf oExistedStdObj Is CStdWord Then
                    '기존 Value가 CStdWord instance인 경우 --> CStdWordCol 생성하고 대체
                    Set oStdWordCol = New CStdWordCol
                    oStdWordCol.Add oExistedStdObj 'Collection에 기존 단어추가
                    oStdWordCol.Add oStdWord 'Collection에 현재 단어 추가

                    m_oStdWordDic.Remove oStdWord.m_s단어논리명 '기존 Key-Value(Instance) 삭제
                    m_oStdWordDic.Add oStdWord.m_s단어논리명, oStdWordCol '새로 Key-Value(Collection) 추가

                ElseIf TypeOf oExistedStdObj Is CStdWordCol Then
                    '기존 Value가 CStdWord Collection인 경우 --> 이미 동음이의어가 있는 경우. 현재 단어를 CStdWordCol에 Add
                    oExistedStdObj.Add oStdWord
                End If
            Else
                '단어 논리명 중복(동음이의어) 불허용
                Set oStdWordTmp = m_oStdWordDic.Item(oStdWord.m_s단어논리명)
                sInfoMsg = "다음의 이유로 실행이 중단되었습니다: " & "단어논리명 중복" & vbLf & _
                           "▶ 중복 항목" & vbLf & _
                               vbTab & "단어논리명: " & oStdWord.m_s단어논리명 & vbLf & _
                               vbTab & "단어물리명: " & oStdWord.m_s단어물리명 & vbLf & _
                           "▶ 기존 항목" & vbLf & _
                               vbTab & "단어논리명: " & oStdWordTmp.m_s단어논리명 & vbLf & _
                               vbTab & "단어물리명: " & oStdWordTmp.m_s단어물리명
                If IsOkToGo(sInfoMsg, "단어 논리명 중복") Then
                    'Resume Next
                Else
                    End
                     Exit For
                End If
            End If
        Else
            If (oStdWord.m_b표준여부 = False) And _
               (oStdWord.m_s표준논리명 = "") Then
                GoTo Skip_단어논리명Build '표준여부 False이고 표준논리명이 없는 경우 Skip
            End If
            '단어논리명 최초 등록
            m_oStdWordDic.Add oStdWord.m_s단어논리명, oStdWord
        End If

Skip_단어논리명Build:

On Error Resume Next
        '단어물리명 기준 Dictionary Build
        If oStdWord.m_b표준여부 = True Then '표준단어일경우만 물리명 기준 Dictionary에 추가(물리명 중복 방지)
            If m_oStdWordDicP.Exists(oStdWord.m_s단어물리명) Then
                oStdWord.m_b물리명중복여부 = True '중복된 물리명이 있는지 체크
                m_oStdWordDicP(oStdWord.m_s단어물리명).m_b물리명중복여부 = True
            End If
            m_oStdWordDicP.Add oStdWord.m_s단어물리명, oStdWord
            If Err <> 0 And aIsAllowDupWordPhysicalName = False Then '단어 물리명 중복 발생
                Set oStdWordTmp = m_oStdWordDicP.Item(oStdWord.m_s단어물리명)
                sInfoMsg = "다음의 이유로 실행이 중단되었습니다: " & "단어물리명 중복: " & vbLf & _
                           "▶ 중복 항목" & vbLf & _
                               vbTab & "단어논리명: " & oStdWord.m_s단어논리명 & vbLf & _
                               vbTab & "단어물리명: " & oStdWord.m_s단어물리명 & vbLf & _
                           "▶ 기존 항목" & vbLf & _
                               vbTab & "단어논리명: " & oStdWordTmp.m_s단어논리명 & vbLf & _
                               vbTab & "단어물리명: " & oStdWordTmp.m_s단어물리명 & vbLf & vbLf & _
                           "중복을 무시하고 계속 진행하시겠습니까?"
    
                If IsOkToGo(sInfoMsg, "단어 물리명 중복") Then
                    Resume Next
                Else
                    End
                    Exit For
                End If
            End If
        End If
On Error GoTo 0

    Next lRow

On Error GoTo 0
End Sub

'논리명 기준의 Item 찾기
Public Function Exists(aWord As String) As Boolean
    Exists = m_oStdWordDic.Exists(aWord)
End Function

'물리명 기준의 Item 찾기
Public Function ExistsP(aWord As String) As Boolean
    ExistsP = m_oStdWordDicP.Exists(aWord)
End Function

Public Function Item(aWord As String) As Object
    Set Item = m_oStdWordDic.Item(aWord)
End Function

Public Function ItemP(aWord As String) As CStdWord
    Set ItemP = m_oStdWordDicP.Item(aWord)
End Function

Public Function Count() As Long
    Count = m_oStdWordDic.Count
End Function
