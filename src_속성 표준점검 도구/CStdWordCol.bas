VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdWordCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'2018-05-13 ǥ�شܾ��� �������Ǿ� ������ ���� �߰�
'CStdWordDic�� �ڷᱸ�� ����
'  ������ --> Key: ǥ�شܾ� ����, Value: CStdWord instance
'  ������ --> Key: ǥ�شܾ� ����, Value: CStdWord instance�� collection (CStdWordCol object)

Public m_oWordCol As Collection 'Key:�ܾ����, Item:CStdWord instance
Private m_lStdIndex As Long '��ǥ��

Private Sub Class_Initialize()
    Set m_oWordCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_oWordCol = Nothing
End Sub

Property Get Items() As Collection
    Set Items = m_oWordCol
End Property

Property Get Count() As Long
    Count = m_oWordCol.Count
End Property

Public Sub Add(aWord As CStdWord)
    m_oWordCol.Add aWord ', aWord.m_s�ܾ����
End Sub

'Public Function Exists(aWordPName As String) As Boolean
'    Exists = False
'
'    On Error GoTo Exit_Exists:
'    m_oWordCol.Item aWordPName
'    Exists = True
'Exit_Exists:
'End Function
'
'Public Function GetWordByPName(aWordPName As String) As CStdWord
'    Set GetWordByPName = Nothing
'
'    On Error GoTo Exit_GetWordByPName:
'    Set GetWordByPName = m_oWordCol.Item(aWordPName)
'
'GetWordByPName:
'End Function

