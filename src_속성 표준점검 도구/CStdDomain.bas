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

Public m_s�����κз��� As String
Public m_s�����γ��� As String
Public m_s�����ι����� As String
Public m_s�����μ��� As String
Public m_s������Ÿ�Ը� As String
Public m_i���� As Integer
Public m_i���� As Integer
Public m_s������Ÿ�Ա��̸� As String

Public Sub SetDomain(a������Ÿ�Ը� As String)
    Dim iPos1 As Integer: iPos1 = 1
    Dim iPos2 As Integer: iPos2 = 1
    Dim iPos3 As Integer: iPos3 = 1

    Me.m_s������Ÿ�Ա��̸� = a������Ÿ�Ը�
    Me.m_i���� = 0
    Me.m_i���� = 0
    iPos1 = InStr(1, a������Ÿ�Ը�, "(")
    If iPos1 = 0 Then ' ��ȣ ���� ���
        Me.m_s������Ÿ�Ը� = a������Ÿ�Ը�
        Exit Sub
    End If
    Me.m_s������Ÿ�Ը� = Mid(a������Ÿ�Ը�, 1, iPos1 - 1)

    iPos2 = InStr(iPos1, a������Ÿ�Ը�, ",")
    If iPos2 = 0 Then ' comma ���� ���
        Me.m_i���� = Mid(a������Ÿ�Ը�, iPos1 + 1, Len(a������Ÿ�Ը�) - iPos1 - 1)
        Exit Sub
    End If
    Me.m_i���� = Mid(a������Ÿ�Ը�, iPos1 + 1, iPos2 - iPos1 - 1)
    Me.m_i���� = Mid(a������Ÿ�Ը�, iPos2 + 1, Len(a������Ÿ�Ը�) - iPos2 - 1)
End Sub

