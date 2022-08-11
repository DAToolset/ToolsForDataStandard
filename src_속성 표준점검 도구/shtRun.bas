VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub chkOnlyForSelectedAttr_Click()
    Selection.Select
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show (vbModal)
End Sub

Private Sub cmdAppendWord_Click()
    �ĺ��ܾ��߰� aAttrBaseRange:=Names("�Ӽ����Base").RefersToRange.Offset(1, 0) _
               , a�ĺ��ܾ�BaseRange:=Names("�ĺ��ܾ���Base").RefersToRange.Offset(1, 0)
End Sub

Private Sub cmdClear_Click()
    If IsOkToGo("ǥ�� ���˰���� �ʱ�ȭ�մϴ�." + vbLf + "��� �����Ͻðڽ��ϱ�?") = False Then Exit Sub
    ClearResult
End Sub

Private Sub cmdConfig_Click()
    shtConfig.Activate
End Sub

Private Sub cmdRefreshStdDic_Click()
    'ǥ�ػ��� ����
    If IsOkToGo("ǥ�ػ����� ���ΰ�ħ�մϴ�." + vbLf + "��� �����Ͻðڽ��ϱ�?") = False Then Exit Sub
    ǥ�ػ������ΰ�ħ Range("ǥ�ػ��������Ͻ�")
End Sub

Private Sub cmdRun_Click()
    Dim sSelectedRunMsg As String, bIsOnlyForSelectedAttr As Boolean, oSelectedAttrRange As Range
    If (chkOnlyForSelectedAttr.Value = True) Or IsShiftKeyDown Then
        sSelectedRunMsg = "������ ���: ������ �Ӽ���"
        bIsOnlyForSelectedAttr = True
        '�ٸ� ���� �����ߴ��� �Ӽ���, DataType/Len �÷� Range ���ϱ�
        Set oSelectedAttrRange = Range("�Ӽ����Base").Offset(1, 0)
        Dim lRowOffset As Long
        lRowOffset = Selection.Row - oSelectedAttrRange.Row
        Set oSelectedAttrRange = oSelectedAttrRange.Offset(lRowOffset, 0).Resize(Selection.Rows.Count, 2)
    Else
        sSelectedRunMsg = "���� ���: ��ü �Ӽ�"
        bIsOnlyForSelectedAttr = False
    End If

    If chkRefreshStdDic.Value = True Then
        If IsOkToGo("ǥ�ػ����� ���ΰ�ħ�ϰ� ǥ�� ������ �����մϴ�." + vbLf + _
                    "(���� ǥ�ػ����� ���)" + vbLf + _
                    sSelectedRunMsg + vbLf + vbLf + _
                    "�ڰ� sheet�� ������ ���͸� �����մϴ١�" + vbLf + vbLf + _
                    "�ð��� �� ������ �ɸ� �� �ֽ��ϴ�." + vbLf + _
                    "��� �����Ͻðڽ��ϱ�?") = False Then
            End
            Exit Sub
        End If
        ǥ�ػ������ΰ�ħ Range("ǥ�ػ��������Ͻ�")
        'Range("ǥ�ػ��������Ͻ�").Value2 = "ǥ�ػ��� �����Ͻ�: " + Format(Now, "yyyy-mm-dd hh:nn:ss")
    Else
        If IsOkToGo("ǥ�� ������ �����մϴ�.(ǥ�ػ��� ����)" + vbLf + _
                    sSelectedRunMsg + vbLf + vbLf + _
                    "�ڰ� sheet�� ������ ���͸� �����մϴ١�" + vbLf + vbLf + _
                    "�ð��� �� ������ �ɸ� �� �ֽ��ϴ�." + vbLf + _
                    "��� �����Ͻðڽ��ϱ�?") = False Then
            End
            Exit Sub
        End If
    End If
    Dim eWordMatchDirection As WordMatchDirection
    If (chkLtoR = False) And (chkRtoL = False) Then
        Call MsgBox("�ܾ����չ����� �����ϼ���.", vbCritical Or vbDefaultButton1, "����")
        Exit Sub
    End If
    If chkLtoR Then eWordMatchDirection = LtoR
    If chkRtoL Then eWordMatchDirection = RtoL
    If (chkLtoR) And (chkRtoL) Then eWordMatchDirection = Both

    ClearResult bIsOnlyForSelectedAttr 'ǥ�����˰�� ���� ����
    ClearAllFilters '�� Sheet�� Filter ����

    Dim eStdDicMatchOption As StdDicMatchOption
    If optWordAndTerm = True Then
        eStdDicMatchOption = WordAndTerm
    ElseIf optWordOnly = True Then
        eStdDicMatchOption = WordOnly
    ElseIf optTermOnly = True Then
        eStdDicMatchOption = TermOnly
    End If

    'ǥ������ Range("B3"), Range("D3"), Range("E3"), eStdDicMatchOption
    ǥ������ aAttrBaseRange:=Range("�Ӽ����Base").Offset(1, 0) _
             , aLNameBaseRange:=Range("ǥ�شܾ��������Base").Offset(1, 0) _
             , aPNameBaseRange:=Range("ǥ�شܾ��������Base").Offset(1, 0) _
             , aStdWordBaseRange:=Sheets("ǥ�شܾ����").Range("B2") _
             , aStdTermBaseRange:=Sheets("ǥ�ؿ�����").Range("B2") _
             , aStdDomainBaseRange:=Sheets("ǥ�ص����λ���").Range("B2") _
             , aStdDicMatchOption:=eStdDicMatchOption _
             , aWordMatchDirection:=eWordMatchDirection _
             , aIsAllowDupWordLogicalName:=chkAllowDupWordLogicalName.Value _
             , aIsAllowDupWordPhysicalName:=chkAllowDupWordPhysicalName.Value _
             , aIsOnlyForSelectedAttr:=bIsOnlyForSelectedAttr _
             , aSelectedAttrRange:=oSelectedAttrRange
End Sub

'ǥ�����˰�� ���� ���� (���� ������ �ʱ�ȭ)
Public Sub ClearResult(Optional aIsOnlyForSelectedAttr As Boolean = False)
    Dim oCurRange As Range: Set oCurRange = Selection 'ActiveCell
    Dim oBaseRange As Range
    Dim lColOffset As Long: lColOffset = 7 'Clear�� �÷��� Offset(= ����-1)
    Dim lRowOffset As Long
    Application.ScreenUpdating = False
    'oBaseRange.Select
    Set oBaseRange = Range("ǥ�شܾ��������Base").Offset(1, 0)
    'Range(oBaseRange, oBaseRange.End(xlDown).Offset(, lColOffset)).ClearContents
    If aIsOnlyForSelectedAttr = False Then '��ü �Ӽ� ���ϰ�� ����
        Range(oBaseRange, oBaseRange.Offset(shtRun.UsedRange.Rows.Count, lColOffset)).ClearContents
    Else '������ �Ӽ� ���˰���� ����
        lRowOffset = oCurRange.Row - oBaseRange.Row
        Set oBaseRange = oBaseRange.Offset(lRowOffset, 0)
        Range(oBaseRange, oBaseRange.Offset(oCurRange.Rows.Count - 1, lColOffset)).ClearContents
    End If

    oCurRange.Select
    Application.ScreenUpdating = True
End Sub

'ǥ������ ������ �� Sheet�� Filter ����
Private Sub ClearAllFilters()
    On Error Resume Next 'Filter�� �����Ǿ� ���� ���� ��� ������ �߻��ϹǷ� Skip
    Sheets("�Ӽ� ǥ������").AutoFilter.ShowAllData
    Sheets("ǥ�شܾ����").AutoFilter.ShowAllData
    Sheets("ǥ�ؿ�����").AutoFilter.ShowAllData
    Sheets("ǥ�ص����λ���").AutoFilter.ShowAllData
    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate()
    ResetControlSize ActiveSheet
End Sub

Public Sub PartialRunTest()
    Dim lRow As Long, oSelRange As Range
    Set oSelRange = Selection
    lRow = oSelRange.Rows.Count
    Debug.Print lRow
    Debug.Print oSelRange.Rows(2).Address
End Sub
