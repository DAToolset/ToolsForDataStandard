VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdBuildConnectionString_Click()
'Connection String 만들기
'참조: Microsoft OLE DB Service Component 1.0 Type Library
On Error GoTo 0

    Dim cn As ADODB.Connection, MSDASCObj As MSDASC.DataLinks, oCurrentRange As Range
    Dim eOrgXlEnableCancelKey As XlEnableCancelKey

'    Set oCurrentRange = ActiveCell
'
'    Set m_oUndoRange = oCurrentRange
'    m_sUndoFormula = m_oUndoRange.Formula

    Set MSDASCObj = New MSDASC.DataLinks

    eOrgXlEnableCancelKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlDisabled
    Set cn = New ADODB.Connection
    'cn.ConnectionString = Range("ConnectionString").Cells.Text
    cn.ConnectionString = Range("ConnectionString").Value2
    If MSDASCObj.PromptEdit(cn) = True Then
        'Range("ConnectionString").Cells.Text = cn.ConnectionString
        'Range("ConnectionString") = cn.ConnectionString
        Range("ConnectionString").Value2 = cn.ConnectionString
    End If

    Set cn = Nothing
    Set MSDASCObj = Nothing

    'Application.OnUndo "Undo build the connection string", "UndoBuildConnectionString"
    Application.EnableCancelKey = eOrgXlEnableCancelKey
End Sub

Private Sub cmdGotoRunSheet_Click()
    shtRun.Activate
End Sub

Private Sub Worksheet_Activate()
    ResetControlSize ActiveSheet
End Sub
