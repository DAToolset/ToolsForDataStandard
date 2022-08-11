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

'2018-05-13 표준단어의 동음이의어 지원을 위해 추가
'CStdWordDic의 자료구조 변경
'  변경전 --> Key: 표준단어 논리명, Value: CStdWord instance
'  변경후 --> Key: 표준단어 논리명, Value: CStdWord instance의 collection (CStdWordCol object)

Public m_oWordCol As Collection 'Key:단어물리명, Item:CStdWord instance
Private m_lStdIndex As Long '비표준

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
    m_oWordCol.Add aWord ', aWord.m_s단어물리명
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

