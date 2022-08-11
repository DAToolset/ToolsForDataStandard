VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'표준단어 Class
Option Explicit

Public m_s단어논리명 As String
Public m_s단어물리명 As String
Public m_s단어영문명 As String
Public m_s단어설명 As String
'Public m_s단어유형 As String
Public m_b표준여부 As Boolean
Public m_b속성분류어여부 As Boolean
Public m_s표준논리명 As String
Public m_s동의어 As String
Public m_s도메인분류명 As String
Public m_oStdWordDic As CStdWordDic
Public m_b논리명중복여부 As Boolean
Public m_b물리명중복여부 As Boolean

Public Sub SetStdWordDic(aStdWordDic As CStdWordDic)
    Set Me.m_oStdWordDic = aStdWordDic
End Sub

Public Function Clone() As CStdWord
    Set Clone = New CStdWord
    With Clone
        .m_s단어논리명 = Me.m_s단어논리명
        .m_s단어물리명 = Me.m_s단어물리명
        .m_s단어영문명 = Me.m_s단어영문명
        .m_s단어설명 = Me.m_s단어설명
        .m_b표준여부 = Me.m_b표준여부
        .m_b속성분류어여부 = Me.m_b속성분류어여부
        .m_s표준논리명 = Me.m_s표준논리명
        .m_s동의어 = Me.m_s동의어
        .m_b논리명중복여부 = Me.m_b논리명중복여부
        .m_b물리명중복여부 = Me.m_b물리명중복여부
    End With
End Function

Public Function GetLastWordChk() As String
    If Me.m_b속성분류어여부 = True Then
        GetLastWordChk = "분류단어"
    Else
        GetLastWordChk = "기본단어"
    End If
End Function

'비표준단어의 표준단어 개체 가져오기
'SW: Standard Word (표준단어)
'NSW: Non-Standard Word (비표준단어)
Public Function GetSWForNSW() As CStdWord
    If Me.m_b표준여부 = True Then
        Set GetSWForNSW = Me '표준단어이면 자기 자신 Return
    Else
        If m_oStdWordDic.Exists(Me.m_s표준논리명) Then
            '비표준단어의 표준논리명이 단어사전에 존재하는 경우
            Set GetSWForNSW = m_oStdWordDic.Item(Me.m_s표준논리명)
        Else
            '비표준단어의 표준논리명이 단어사전에 없는 경우
            Set GetSWForNSW = New CStdWord
            GetSWForNSW.m_s단어논리명 = "<" + Me.m_s단어논리명 + ": 표준단어 없음>"
            GetSWForNSW.m_s단어물리명 = "<" + Me.m_s단어물리명 + ": 표준단어 없음>"
        End If
    End If
End Function

'생성자: 변수 초기화
Private Sub Class_Initialize()
    m_b표준여부 = False
    m_b속성분류어여부 = False
    m_b논리명중복여부 = False
    m_b물리명중복여부 = False
End Sub
