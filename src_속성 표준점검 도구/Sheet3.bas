VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdRun_Click()
    Dim aAttrBaseRange As Range: Set aAttrBaseRange = Range("C4")
    Dim aStdWordBaseRange As Range: Set aStdWordBaseRange = Sheets("ǥ�شܾ����").Range("B2")
    Dim aIsAllowDupWordLogicalName As Boolean: aIsAllowDupWordLogicalName = True
    Dim aIsAllowDupWordPhysicalName As Boolean: aIsAllowDupWordPhysicalName = True
    Dim aIsOnlyForSelectedAttr As Boolean: aIsOnlyForSelectedAttr = False
    Dim aSelectedAttrRange As Range

    Dim oStdWordDic As CStdWordDic '�ܾ����
    Dim oStdTermDic As CStdTermDic '������
    Dim oStdDomainDic As CStdDomainDic '�����λ���
    Dim vInRngArr As Variant, vOutRngArr As Variant, lRow As Long, lLastRow As Long, oOutRange As Range

'    Dim eStdDicMatchOption As StdDicMatchOption 'ǥ�ػ��� ã�� �ɼ�
'    Dim eWordMatchDirection As WordMatchDirection '�ܾ����� ����

    '�ܾ���� �ҷ����� ---------------------------------------------------------------------
    Set oStdWordDic = New CStdWordDic: oStdWordDic.m_eStdDicMatchOption = WordAndTerm
    oStdWordDic.Load aStdWordBaseRange, aIsAllowDupWordLogicalName, aIsAllowDupWordPhysicalName
    If oStdWordDic.Count = 0 Then
        MsgBox "�ܾ������ ��� �ֽ��ϴ�." + vbLf + "ǥ�� ������ �����մϴ�", vbCritical, "����"
        Exit Sub
    End If

    Dim oAttrRange As Range, lRowOffset As Long
    Dim sColName As String, sAttrName As String, iLenAttrName As Integer, sAttrDataTypeSize As String
    lRowOffset = 0
    '�Է¹��� �о Variant array�� ���
    If aAttrBaseRange.Offset(1, 0).Value2 = "" Then 'ù �ุ �����Ͱ� �ִ� ���
        vInRngArr = aAttrBaseRange.Resize(, 2).Value2 '�д� ����: 2�� �÷�
    Else
        If Not aIsOnlyForSelectedAttr Then '������ �Ӽ��� ���� �ʴ� ��� --> ��ü �б�
            vInRngArr = Range(aAttrBaseRange, aAttrBaseRange.End(xlDown)).Resize(, 2).Value2 '�д� ����: 2�� �÷�
        Else '������ �Ӽ��� �д� ���
'            lRowOffset = aSelectedAttrRange.Row - aLNameBaseRange.Row
'            vInRngArr = aSelectedAttrRange.Value2
        End If
    End If
    Set oOutRange = Range("F4")

    '------------------------------------------------------------------------------------------
    '������ �������� ���� ã��
    Dim aColNamePart() As String, lColNamePartIdx As Long, oColNamePartAsStdWord As Collection
    Dim oStdWordObj As Object, oStdWord As CStdWord, sLName As String
    For lRow = LBound(vInRngArr) To UBound(vInRngArr)
        sColName = vInRngArr(lRow, 1)
        If Trim(sColName) = "" Then GoTo SkipBlank '������ �Ӽ����� ����ִ� ��� Skip
        aColNamePart = Split(sColName, "_")
        Set oColNamePartAsStdWord = Nothing
        Set oColNamePartAsStdWord = New Collection
        sLName = ""
        For lColNamePartIdx = LBound(aColNamePart) To UBound(aColNamePart)
            Set oStdWordObj = oStdWordDic.ItemP(aColNamePart(lColNamePartIdx))
            oColNamePartAsStdWord.Add oStdWordObj
            If TypeOf oStdWordObj Is CStdWord Then
                Set oStdWord = oStdWordObj
                sLName = sLName + IIf(sLName = "", "", "_") + oStdWord.m_s�ܾ����
            End If
            oOutRange.Offset(lRow - 1, 0).Value2 = sLName
        Next lColNamePartIdx
SkipBlank:
    Next lRow
    '------------------------------------------------------------------------------------------
End Sub
