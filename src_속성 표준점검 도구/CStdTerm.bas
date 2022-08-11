VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ǥ�ؿ�� Class
Option Explicit
'Option Base 1

Public m_s������ As String
Public m_s�ܾ�������� As String
Public m_s������ As String
Public m_s���� As String
Public m_s�����γ��� As String
Public m_s������Ÿ�Ը� As String
Public m_i���� As Integer
Public m_i���� As Integer
Public m_s���Ǿ��� As String

Private m_oStdWordDicP As Dictionary

Public m_sSorted�ܾ�������� As String

'�߰� ��� �Ӽ�
'Public m_sǥ�شܾ����� As String
Public m_s������Ÿ�Ա��̸� As String

Public Sub SetData(aStdWordDic As Dictionary)
    Set m_oStdWordDicP = aStdWordDic
    'If m_s�ܾ�������� = "" Then Setǥ�شܾ����� ' oWordDicMap
    If m_s������Ÿ�Ա��̸� = "" Then _
        m_s������Ÿ�Ա��̸� = GetDataTypeStr(m_s������Ÿ�Ը�, m_i����, m_i����)
'    Select Case m_s������Ÿ�Ը�
'        Case "DATE", "BLOB", "CLOB", "LONG", "TIMESTAMP"
'            m_s������Ÿ�Ա��̸� = m_s������Ÿ�Ը�
'        Case "CHAR", "VARCHAR", "VARCHAR2"
'            m_s������Ÿ�Ա��̸� = m_s������Ÿ�Ը� + "(" + m_s���� + ")"
'        Case "NUMBER"
'            If m_s���� = "" Then
'                m_s������Ÿ�Ա��̸� = m_s������Ÿ�Ը� + "(" + m_s���� + ")"
'            Else
'                m_s������Ÿ�Ա��̸� = m_s������Ÿ�Ը� + "(" + m_s���� + "," + m_s���� + ")"
'            End If
'    End Select

    '2021-02-07 �߰�(�ܾ���� ���� �ߺ����θ� �Ǻ��ϱ� ���� �ܾ������ ����)
    If m_s�ܾ�������� <> "" Then
        Dim a�ܾ����() As String
        ' If m_s�ܾ�������� = "����_����_�ڵ�" Then Stop
        a�ܾ���� = Split(m_s�ܾ��������, "_", , vbTextCompare)
        Call QuickSort(a�ܾ����, LBound(a�ܾ����), UBound(a�ܾ����))
        m_sSorted�ܾ�������� = Join(a�ܾ����, "") 'Join(a�ܾ����, "_") '���վ ����� ������ ���� ������������ �Ǻ��Ϸ��� "_" ������ ���� join�� ���ڿ� ���
    End If
End Sub

Public Sub Setǥ�شܾ�����() '(ByRef oWordDicMap As Dictionary)
    Const C_NOT_EXISTS_WORD = "(NOT_EXISTS)"
    Dim a������() As String, a����() As String, i As Long, oStdWord As CStdWord
    a������ = Split(m_s������, "_")
    ReDim a����(UBound(a������))
    For i = 0 To UBound(a����)
        If m_oStdWordDicP.Exists(a������(i)) Then
            Set oStdWord = m_oStdWordDicP.Item(a������(i))
            a����(i) = oStdWord.m_s�ܾ����
            m_s�ܾ�������� = m_s�ܾ�������� + a����(i) + "_"
        Else
            m_s�ܾ�������� = m_s�ܾ�������� + C_NOT_EXISTS_WORD + "_"
        End If
    Next i
    m_s�ܾ�������� = Mid(m_s�ܾ��������, 1, Len(m_s�ܾ��������) - 1)
End Sub

Public Function Clone() As CStdTerm
    Set Clone = New CStdTerm
    With Clone
        .m_s������ = Me.m_s������
        .m_s�ܾ�������� = Me.m_s�ܾ��������
        .m_s������ = Me.m_s������
        .m_s���� = Me.m_s����
        .m_s�����γ��� = Me.m_s�����γ���
        .m_s������Ÿ�Ը� = Me.m_s������Ÿ�Ը�
        .m_i���� = Me.m_i����
        .m_i���� = Me.m_i����
        .m_s���Ǿ��� = Me.m_s���Ǿ���
    End With
End Function

'Public Sub SetTerm(a�Ӽ��� As String, a������Ÿ�Ը� As String)
'    Dim iPos1 As Integer: iPos1 = 1
'    Dim iPos2 As Integer: iPos2 = 1
'    Dim iPos3 As Integer: iPos3 = 1
'
'    Me.m_s������ = a�Ӽ���
'    Me.m_s������Ÿ�Ա��̸� = a������Ÿ�Ը�
'    Me.m_s���� = "0"
'    Me.m_s���� = "0"
'    iPos1 = InStr(1, a������Ÿ�Ը�, "(")
'    If iPos1 = 0 Then ' ��ȣ ���� ���
'        Me.m_s������Ÿ�Ը� = a������Ÿ�Ը�
'        Exit Sub
'    End If
'    Me.m_s������Ÿ�Ը� = Mid(a������Ÿ�Ը�, 1, iPos1 - 1)
'
'    iPos2 = InStr(iPos1, a������Ÿ�Ը�, ",")
'    If iPos2 = 0 Then ' comma ���� ���
'        Me.m_s���� = Mid(a������Ÿ�Ը�, iPos1 + 1, Len(a������Ÿ�Ը�) - iPos1 - 1)
'        Exit Sub
'    End If
'    Me.m_s���� = Mid(a������Ÿ�Ը�, iPos1 + 1, iPos2 - iPos1 - 1)
'    Me.m_s���� = Mid(a������Ÿ�Ը�, iPos2 + 1, Len(a������Ÿ�Ը�) - iPos2 - 1)
'End Sub

