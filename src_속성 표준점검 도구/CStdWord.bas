VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ǥ�شܾ� Class
Option Explicit

Public m_s�ܾ���� As String
Public m_s�ܾ���� As String
Public m_s�ܾ���� As String
Public m_s�ܾ�� As String
'Public m_s�ܾ����� As String
Public m_bǥ�ؿ��� As Boolean
Public m_b�Ӽ��з���� As Boolean
Public m_sǥ�س��� As String
Public m_s���Ǿ� As String
Public m_s�����κз��� As String
Public m_oStdWordDic As CStdWordDic
Public m_b�����ߺ����� As Boolean
Public m_b�������ߺ����� As Boolean

Public Sub SetStdWordDic(aStdWordDic As CStdWordDic)
    Set Me.m_oStdWordDic = aStdWordDic
End Sub

Public Function Clone() As CStdWord
    Set Clone = New CStdWord
    With Clone
        .m_s�ܾ���� = Me.m_s�ܾ����
        .m_s�ܾ���� = Me.m_s�ܾ����
        .m_s�ܾ���� = Me.m_s�ܾ����
        .m_s�ܾ�� = Me.m_s�ܾ��
        .m_bǥ�ؿ��� = Me.m_bǥ�ؿ���
        .m_b�Ӽ��з���� = Me.m_b�Ӽ��з����
        .m_sǥ�س��� = Me.m_sǥ�س���
        .m_s���Ǿ� = Me.m_s���Ǿ�
        .m_b�����ߺ����� = Me.m_b�����ߺ�����
        .m_b�������ߺ����� = Me.m_b�������ߺ�����
    End With
End Function

Public Function GetLastWordChk() As String
    If Me.m_b�Ӽ��з���� = True Then
        GetLastWordChk = "�з��ܾ�"
    Else
        GetLastWordChk = "�⺻�ܾ�"
    End If
End Function

'��ǥ�شܾ��� ǥ�شܾ� ��ü ��������
'SW: Standard Word (ǥ�شܾ�)
'NSW: Non-Standard Word (��ǥ�شܾ�)
Public Function GetSWForNSW() As CStdWord
    If Me.m_bǥ�ؿ��� = True Then
        Set GetSWForNSW = Me 'ǥ�شܾ��̸� �ڱ� �ڽ� Return
    Else
        If m_oStdWordDic.Exists(Me.m_sǥ�س���) Then
            '��ǥ�شܾ��� ǥ�س����� �ܾ������ �����ϴ� ���
            Set GetSWForNSW = m_oStdWordDic.Item(Me.m_sǥ�س���)
        Else
            '��ǥ�شܾ��� ǥ�س����� �ܾ������ ���� ���
            Set GetSWForNSW = New CStdWord
            GetSWForNSW.m_s�ܾ���� = "<" + Me.m_s�ܾ���� + ": ǥ�شܾ� ����>"
            GetSWForNSW.m_s�ܾ���� = "<" + Me.m_s�ܾ���� + ": ǥ�شܾ� ����>"
        End If
    End If
End Function

'������: ���� �ʱ�ȭ
Private Sub Class_Initialize()
    m_bǥ�ؿ��� = False
    m_b�Ӽ��з���� = False
    m_b�����ߺ����� = False
    m_b�������ߺ����� = False
End Sub
