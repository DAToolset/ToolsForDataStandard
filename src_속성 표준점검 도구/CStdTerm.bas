VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'표준용어 Class
Option Explicit
'Option Base 1

Public m_s용어논리명 As String
Public m_s단어논리명조합 As String
Public m_s용어물리명 As String
Public m_s용어설명 As String
Public m_s도메인논리명 As String
Public m_s데이터타입명 As String
Public m_i길이 As Integer
Public m_i정도 As Integer
Public m_s정의업무 As String

Private m_oStdWordDicP As Dictionary

Public m_sSorted단어논리명조합 As String

'추가 멤버 속성
'Public m_s표준단어조합 As String
Public m_s데이터타입길이명 As String

Public Sub SetData(aStdWordDic As Dictionary)
    Set m_oStdWordDicP = aStdWordDic
    'If m_s단어논리명조합 = "" Then Set표준단어조합 ' oWordDicMap
    If m_s데이터타입길이명 = "" Then _
        m_s데이터타입길이명 = GetDataTypeStr(m_s데이터타입명, m_i길이, m_i정도)
'    Select Case m_s데이터타입명
'        Case "DATE", "BLOB", "CLOB", "LONG", "TIMESTAMP"
'            m_s데이터타입길이명 = m_s데이터타입명
'        Case "CHAR", "VARCHAR", "VARCHAR2"
'            m_s데이터타입길이명 = m_s데이터타입명 + "(" + m_s길이 + ")"
'        Case "NUMBER"
'            If m_s정도 = "" Then
'                m_s데이터타입길이명 = m_s데이터타입명 + "(" + m_s길이 + ")"
'            Else
'                m_s데이터타입길이명 = m_s데이터타입명 + "(" + m_s길이 + "," + m_s정도 + ")"
'            End If
'    End Select

    '2021-02-07 추가(단어논리명 조합 중복여부를 판별하기 위해 단어논리명을 정렬)
    If m_s단어논리명조합 <> "" Then
        Dim a단어논리명() As String
        ' If m_s단어논리명조합 = "휴점_지역_코드" Then Stop
        a단어논리명 = Split(m_s단어논리명조합, "_", , vbTextCompare)
        Call QuickSort(a단어논리명, LBound(a단어논리명), UBound(a단어논리명))
        m_sSorted단어논리명조합 = Join(a단어논리명, "") 'Join(a단어논리명, "_") '복합어를 사용한 경우까지 동일 논리명조합으로 판별하려면 "_" 구분자 없이 join한 문자열 사용
    End If
End Sub

Public Sub Set표준단어조합() '(ByRef oWordDicMap As Dictionary)
    Const C_NOT_EXISTS_WORD = "(NOT_EXISTS)"
    Dim a물리명() As String, a논리명() As String, i As Long, oStdWord As CStdWord
    a물리명 = Split(m_s용어물리명, "_")
    ReDim a논리명(UBound(a물리명))
    For i = 0 To UBound(a논리명)
        If m_oStdWordDicP.Exists(a물리명(i)) Then
            Set oStdWord = m_oStdWordDicP.Item(a물리명(i))
            a논리명(i) = oStdWord.m_s단어논리명
            m_s단어논리명조합 = m_s단어논리명조합 + a논리명(i) + "_"
        Else
            m_s단어논리명조합 = m_s단어논리명조합 + C_NOT_EXISTS_WORD + "_"
        End If
    Next i
    m_s단어논리명조합 = Mid(m_s단어논리명조합, 1, Len(m_s단어논리명조합) - 1)
End Sub

Public Function Clone() As CStdTerm
    Set Clone = New CStdTerm
    With Clone
        .m_s용어논리명 = Me.m_s용어논리명
        .m_s단어논리명조합 = Me.m_s단어논리명조합
        .m_s용어물리명 = Me.m_s용어물리명
        .m_s용어설명 = Me.m_s용어설명
        .m_s도메인논리명 = Me.m_s도메인논리명
        .m_s데이터타입명 = Me.m_s데이터타입명
        .m_i길이 = Me.m_i길이
        .m_i정도 = Me.m_i정도
        .m_s정의업무 = Me.m_s정의업무
    End With
End Function

'Public Sub SetTerm(a속성명 As String, a데이터타입명 As String)
'    Dim iPos1 As Integer: iPos1 = 1
'    Dim iPos2 As Integer: iPos2 = 1
'    Dim iPos3 As Integer: iPos3 = 1
'
'    Me.m_s용어논리명 = a속성명
'    Me.m_s데이터타입길이명 = a데이터타입명
'    Me.m_s길이 = "0"
'    Me.m_s정도 = "0"
'    iPos1 = InStr(1, a데이터타입명, "(")
'    If iPos1 = 0 Then ' 괄호 없는 경우
'        Me.m_s데이터타입명 = a데이터타입명
'        Exit Sub
'    End If
'    Me.m_s데이터타입명 = Mid(a데이터타입명, 1, iPos1 - 1)
'
'    iPos2 = InStr(iPos1, a데이터타입명, ",")
'    If iPos2 = 0 Then ' comma 없는 경우
'        Me.m_s길이 = Mid(a데이터타입명, iPos1 + 1, Len(a데이터타입명) - iPos1 - 1)
'        Exit Sub
'    End If
'    Me.m_s길이 = Mid(a데이터타입명, iPos1 + 1, iPos2 - iPos1 - 1)
'    Me.m_s정도 = Mid(a데이터타입명, iPos2 + 1, Len(a데이터타입명) - iPos2 - 1)
'End Sub

