VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sConnectionName As String
Private m_sConnectionString As String
'Private m_oCon As Object
'Private m_oRs As Object
Private m_oCon As ADODB.Connection
Private m_oRs As ADODB.Recordset
Private Const sHeaderAddr As String = "A1"
Private Const sDataAddr As String = "A2"

Public Sub InitProperty(aConnectionName As String, aConnectionString As String)
    m_sConnectionName = aConnectionName
    m_sConnectionString = aConnectionString
    'Set m_oCon = CreateObject("ADODB.Connection")
    'Set m_oRs = CreateObject("ADODB.Recordset")

    Set m_oCon = New ADODB.Connection
    Set m_oRs = New ADODB.Recordset
End Sub

Public Function IsOpen() As Boolean
    IsOpen = False
    If Not m_oCon Is Nothing And _
       m_oCon.State <> 0 Then 'adStateClosed Then
       IsOpen = True
    End If
End Function

Public Sub DoConnect()
    If IsOpen Then Exit Sub
    m_oCon.ConnectionTimeout = 5
    m_oCon.Open m_sConnectionString
End Sub

Public Sub PopulateQueryResult(aSql As String, aSht As Worksheet)
    Dim iRow As Long, iCol As Long
    DoConnect
    m_oRs.Open aSql, m_oCon, 0 'adOpenForwardOnly

    iRow = 0

'    'Column Header 설정
'    For iCol = 0 To m_oRs.Fields.Count - 1
'        aSht.Range(sHeaderAddr).Offset(iRow, iCol).Value = m_oRs.Fields(iCol).Name
'    Next iCol

    'Data 가져오기
    aSht.Range(sDataAddr).CopyFromRecordset m_oRs

    m_oRs.Close

'    DoLog "Formating..."
'    DoResultFormatting aSht
'    DoLog "Formating has finished."
End Sub

Private Sub Class_Terminate()
    If Not m_oRs Is Nothing Then
        If m_oRs.State <> adStateClosed Then m_oRs.Close
        Set m_oRs = Nothing
    End If

    If Not m_oCon Is Nothing Then
        If m_oCon.State <> adStateClosed Then m_oCon.Close
        Set m_oCon = Nothing
    End If

End Sub
