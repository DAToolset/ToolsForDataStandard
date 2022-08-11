VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdDomainDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_oStdDomainDic As Dictionary 'Key: 도메인분류명, Value: CStdDomain의 Collection
Public m_oStdDomainDicT As Dictionary 'Key: 도메인분류명, Value: Dictionary(Key:m_s데이터타입길이명, Value: CStdDomain의 Collection)
Public m_oStdWordDic As CStdWordDic '단어사전(속성분류어의 도메인분류명을 확인하는 용도)
Public m_eStdDicMatchOption As StdDicMatchOption

Private Sub Class_Initialize()
    Set m_oStdDomainDic = New Dictionary
    Set m_oStdDomainDicT = New Dictionary
End Sub

Private Sub Class_Terminate()
    m_oStdDomainDic.RemoveAll
    Set m_oStdDomainDic = Nothing

    m_oStdDomainDicT.RemoveAll
    Set m_oStdDomainDicT = Nothing
End Sub

Public Sub SetStdWordDic(aStdWordDic As CStdWordDic)
    Set m_oStdWordDic = aStdWordDic
End Sub

Public Sub Add(aStdDomain As CStdDomain)
    Dim oStdDomainCol As Collection
    If m_oStdDomainDic.Exists(aStdDomain.m_s도메인분류명) Then
        Set oStdDomainCol = m_oStdDomainDic(aStdDomain.m_s도메인분류명)
    Else
        Set oStdDomainCol = New Collection
        m_oStdDomainDic.Add aStdDomain.m_s도메인분류명, oStdDomainCol
    End If
    oStdDomainCol.Add aStdDomain
End Sub

Public Function GetDomainCollection(a도메인분류명 As String) As Collection
    If m_oStdDomainDic.Exists(a도메인분류명) Then
        Set GetDomainCollection = m_oStdDomainDic(a도메인분류명)
    End If
End Function

Public Sub Load(aBaseRange As Range)
    Dim oStdDomain As CStdDomain
    Dim lRow As Long

    '목록에 아무 값이 없는 경우 exit
    If Trim(aBaseRange.Offset(1, 0)) = "" Then Exit Sub

    Dim vRngArr As Variant
    vRngArr = Range(aBaseRange, aBaseRange.End(xlDown)).Resize(, 8).Value2 '읽는 범위: 8개 컬럼

    For lRow = LBound(vRngArr) To UBound(vRngArr)
        Set oStdDomain = New CStdDomain
        With oStdDomain
            .m_s도메인분류명 = vRngArr(lRow, 1)
            .m_s도메인논리명 = vRngArr(lRow, 2)
            .m_s도메인물리명 = vRngArr(lRow, 3)
            .m_s도메인설명 = vRngArr(lRow, 4)
            .m_s데이터타입명 = vRngArr(lRow, 5)
            .m_i길이 = vRngArr(lRow, 6)
            .m_i정도 = vRngArr(lRow, 7)
            .m_s데이터타입길이명 = vRngArr(lRow, 8)
            '.m_s데이터타입길이명 = GetDataTypeStr(.m_s데이터타입명, .m_i길이, .m_i정도)
        End With

        Me.Add oStdDomain
    Next lRow

    '데이터타입길이명 검색용 Dictionary build
    Dim oKey As Variant, s도메인분류명 As String
    Dim oStdDomainCollection As Collection, oTypeDic As Dictionary
    For Each oKey In m_oStdDomainDic.Keys
        s도메인분류명 = oKey
        'Set oStdDomain = m_oStdDomainDic(s도메인분류명)
        Set oStdDomainCollection = m_oStdDomainDic(s도메인분류명)
        For Each oStdDomain In oStdDomainCollection
            If m_oStdDomainDicT.Exists(s도메인분류명) Then
                Set oTypeDic = m_oStdDomainDicT(s도메인분류명)
            Else
                Set oTypeDic = New Dictionary
                m_oStdDomainDicT.Add s도메인분류명, oTypeDic
            End If
On Error Resume Next '도메인분류명 내의 데이터타입길이명이 중복되도 무시
            oTypeDic.Add oStdDomain.m_s데이터타입길이명, oStdDomain
On Error GoTo 0
        Next
    Next
End Sub

'속성데이터타입길이명과 표준용어의 데이터타입길이명을 비교한 결과 return
'Public Function GetCheckAttrDataType(a속성종결어 As String, _
'       a속성데이터타입길이명 As String, a표준용어데이터타입길이명 As String)
Public Function GetCheckAttrDataType(a속성종결어 As String, a속성종결어StdWord As CStdWord, _
        a속성데이터타입길이명 As String, a표준용어데이터타입길이명 As String, _
        Optional aStdWord As CStdWord) As String

    Dim sResult As String, oStdDomain As CStdDomain, oStdDomainTerm As CStdDomain, oStdDomainAtt As CStdDomain
    Dim s속성종결어_도메인분류명 As String, oStdWord As CStdWord, b속성종결어_도메인분류존재여부 As Boolean

    If a속성데이터타입길이명 = "" Then
        sResult = "속성 Data Type 지정 필요"
        GetCheckAttrDataType = sResult
        Exit Function
    End If

    Set oStdDomainAtt = New CStdDomain '속성 데이터타입(비교용)
    oStdDomainAtt.SetDomain a속성데이터타입길이명

    If a표준용어데이터타입길이명 > "" Then
        '------------------------------------------------------------------------------------------
        '표준용어로 매칭된 경우
        Set oStdDomainTerm = New CStdDomain '표준용어데이터타입(비교용)
        oStdDomainTerm.SetDomain a표준용어데이터타입길이명

        sResult = GetCompareResult("표준용어", oStdDomainAtt, oStdDomainTerm)
    Else
        '------------------------------------------------------------------------------------------
        '표준단어로 매칭된 경우

        sResult = "도메인 점검 결과"
        '속성종결어로 도메인분류 찾기
        s속성종결어_도메인분류명 = a속성종결어
        If Not a속성종결어StdWord Is Nothing Then
            s속성종결어_도메인분류명 = a속성종결어StdWord.m_s도메인분류명
        Else
            sResult = sResult + vbLf + "속성종결어: 표준단어 없음(" & a속성종결어 & ")"
        End If

        Set oStdDomain = GetStdDomain(a도메인분류명:=s속성종결어_도메인분류명 _
                                , a데이터타입길이명:=a속성데이터타입길이명 _
                                , a도메인분류명존재여부:=b속성종결어_도메인분류존재여부)

        If b속성종결어_도메인분류존재여부 = False Then
            sResult = sResult + vbLf + "속성종결어:도메인분류 없음"
        End If

        If Not oStdDomain Is Nothing Then
            '속성 분류어에 지정된 데이터타입 중 속성 데이터 타입이 존재함
            sResult = GetCompareResult("도메인", oStdDomainAtt, oStdDomain)
        Else
            sResult = sResult + vbLf + "도메인 추가 필요 <" + _
                      IIf(s속성종결어_도메인분류명 > "", s속성종결어_도메인분류명, a속성종결어) + _
                      ">: " + a속성데이터타입길이명
        End If
    End If
    '----------------------------------------------------------------------------------------------

    GetCheckAttrDataType = sResult
End Function

Public Function GetStdDomain(a도메인분류명 As String, _
        a데이터타입길이명 As String, a도메인분류명존재여부 As Boolean) As CStdDomain
    Dim oResult As CStdDomain, oTypeDic As Dictionary
    a도메인분류명존재여부 = False
    If m_oStdDomainDicT.Exists(a도메인분류명) Then
        a도메인분류명존재여부 = True
        Set oTypeDic = m_oStdDomainDicT(a도메인분류명)
        If oTypeDic.Exists(a데이터타입길이명) Then Set oResult = oTypeDic(a데이터타입길이명)
    End If
    Set GetStdDomain = oResult
End Function

Public Function ExistsDomainDataType(a도메인분류명 As String, _
        a데이터타입길이명 As String) As Boolean
    Dim bResult As Boolean, oTypeDic As Dictionary
    bResult = False
    If m_oStdDomainDicT.Exists(a도메인분류명) Then
        Set oTypeDic = m_oStdDomainDicT(a도메인분류명)
        If oTypeDic.Exists(a데이터타입길이명) Then bResult = True
    End If
    ExistsDomainDataType = bResult
End Function

'두 Domain의 TypeSize 비교결과 return
'aDomainAtt: 비교기준 Domain (속성지정 Type/Size)
'aDomainTgt: 비교대상 Domain (용어 또는 속성분류어 Type/Size)
Public Function GetCompareResult(aCompareType As String, _
            aDomainAtt As CStdDomain, aDomainTgt As CStdDomain) As String
    Dim sResult As String
    sResult = aCompareType + " Type/Size 비교 결과"
    If aDomainAtt.m_s데이터타입명 <> aDomainTgt.m_s데이터타입명 Then
        sResult = sResult + vbLf + "타입 불일치"
    End If
    If aDomainAtt.m_i길이 <> aDomainTgt.m_i길이 Then
        sResult = sResult + vbLf + "길이 불일치"
        If aDomainAtt.m_i길이 > aDomainTgt.m_i길이 Then '속성의 Size가 더 큰 경우(도메인 추가 또는 속성 size 조정)
            sResult = sResult + "(감소! 도메인 추가 또는 속성 Size 조정 필요)"
        ElseIf aDomainAtt.m_i길이 < aDomainTgt.m_i길이 Then '비교 대상 Domain Size가 더 큰 경우(대부분은 문제없음)
            sResult = sResult + "(증가 확인)"
        End If
    ElseIf aDomainAtt.m_i정도 <> aDomainTgt.m_i정도 Then
        sResult = sResult + vbLf + "소수점 길이 불일치"
        If aDomainAtt.m_i정도 > aDomainTgt.m_i정도 Then '속성의 Size가 더 큰 경우(도메인 추가 또는 속성 size 조정)
            sResult = sResult + "(감소! 도메인 추가 또는 속성 Size 조정 필요)"
        ElseIf aDomainAtt.m_i정도 < aDomainTgt.m_i정도 Then '비교 대상 Domain Size가 더 큰 경우(대부분은 문제없음)
            sResult = sResult + "(증가 확인)"
  End If
    Else
  sResult = aCompareType + " Type/Size 일치"
    End If
    GetCompareResult = sResult
End Function


''두 Domain의 TypeSize 비교결과 return
''aDomainAtt: 비교기준 Domain (속성지정 Type/Size)
''aDomainTgt: 비교대상 Domain (용어 또는 속성분류어 Type/Size)
'Public Function GetCompareResult(aCompareType As String, _
'            aDomainAtt As CStdDomain, aDomainTgt As CStdDomain) As String
'    Dim sResult As String
'    sResult = aCompareType + " Type/Size 비교 결과"
'    If aDomainAtt.m_s데이터타입명 <> aDomainTgt.m_s데이터타입명 Then
'        sResult = sResult + vbLf + "타입 불일치"
'    ElseIf aDomainAtt.m_i길이 <> aDomainTgt.m_i길이 Then
'        sResult = sResult + vbLf + "길이 불일치"
'        If aDomainAtt.m_i길이 > aDomainTgt.m_i길이 Then '속성의 Size가 더 큰 경우(도메인 추가 또는 속성 size 조정)
'            sResult = sResult + "(도메인 추가 또는 속성 Size 조정 필요)"
''        Else '비교 대상 Domain Size가 더 큰 경우(대부분은 문제없음)
''            sResult = sResult + ""
'        End If
'    ElseIf aDomainAtt.m_i정도 <> aDomainTgt.m_i정도 Then
'        sResult = sResult + vbLf + "소수점 길이 불일치"
'        If aDomainAtt.m_i정도 > aDomainTgt.m_i정도 Then '속성의 Size가 더 큰 경우(도메인 추가 또는 속성 size 조정)
'            sResult = sResult + "(도메인 추가 또는 속성 Size 조정 필요)"
''        Else '비교 대상 Domain Size가 더 큰 경우(대부분은 문제없음)
''            sResult = sResult + ""
'  End If
'    Else
'  sResult = aCompareType + " Type/Size 일치"
'    End If
'    GetCompareResult = sResult
'End Function
