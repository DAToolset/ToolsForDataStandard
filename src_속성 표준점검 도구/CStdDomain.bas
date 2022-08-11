VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_s도메인분류명 As String
Public m_s도메인논리명 As String
Public m_s도메인물리명 As String
Public m_s도메인설명 As String
Public m_s데이터타입명 As String
Public m_i길이 As Integer
Public m_i정도 As Integer
Public m_s데이터타입길이명 As String

Public Sub SetDomain(a데이터타입명 As String)
    Dim iPos1 As Integer: iPos1 = 1
    Dim iPos2 As Integer: iPos2 = 1
    Dim iPos3 As Integer: iPos3 = 1

    Me.m_s데이터타입길이명 = a데이터타입명
    Me.m_i길이 = 0
    Me.m_i정도 = 0
    iPos1 = InStr(1, a데이터타입명, "(")
    If iPos1 = 0 Then ' 괄호 없는 경우
        Me.m_s데이터타입명 = a데이터타입명
        Exit Sub
    End If
    Me.m_s데이터타입명 = Mid(a데이터타입명, 1, iPos1 - 1)

    iPos2 = InStr(iPos1, a데이터타입명, ",")
    If iPos2 = 0 Then ' comma 없는 경우
        Me.m_i길이 = Mid(a데이터타입명, iPos1 + 1, Len(a데이터타입명) - iPos1 - 1)
        Exit Sub
    End If
    Me.m_i길이 = Mid(a데이터타입명, iPos1 + 1, iPos2 - iPos1 - 1)
    Me.m_i정도 = Mid(a데이터타입명, iPos2 + 1, Len(a데이터타입명) - iPos2 - 1)
End Sub

