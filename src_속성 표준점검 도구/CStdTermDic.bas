VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdTermDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ǥ�ؿ�� Dictionary Class
Option Explicit

Public m_oStdTermDic As Dictionary 'Key: ������, Value: CStdTerm instance
Private m_oStdWordDicP As Dictionary '�ܾ����(Key:�ܾ����, Value: CStdWord instance)
Public m_o��������Dic As Dictionary '����� �������� ���ϼ� �Ǻ��� ���� dictionary(Key:�ܾ��������, Value: Collection of CStdTerm)
Public m_eStdDicMatchOption As StdDicMatchOption
Public m_s�����ߺ����˰�� As String

Private Sub Class_Initialize()
    Set m_oStdTermDic = New Dictionary
    Set m_o��������Dic = New Dictionary
End Sub

Private Sub Class_Terminate()
    m_oStdTermDic.RemoveAll: Set m_oStdTermDic = Nothing
    m_o��������Dic.RemoveAll: Set m_o��������Dic = Nothing
End Sub

Public Sub SetStdWordDicP(aStdWordDicP As Dictionary)
    Set m_oStdWordDicP = aStdWordDicP
End Sub

'������ Loading(By Variant Array)
Public Sub Load(aBaseRange As Range)
    '�ܾ �����ϴ� ��� ������ Loading���� ����
    If m_eStdDicMatchOption = WordOnly Then Exit Sub

    '��Ͽ� �ƹ� ���� ���� ��� exit
    If Trim(aBaseRange.Offset(1, 0)) = "" Then Exit Sub

    Dim vRngArr As Variant
    vRngArr = Range(aBaseRange, aBaseRange.End(xlDown)).Resize(, 11).Value2 '�д� ����: 10�� �÷�
    Dim lRow As Long, sInfoMsg As String

    Dim oStdTerm As CStdTerm, oStdTermTmp As CStdTerm, oStdTermCol As Collection
    For lRow = LBound(vRngArr) To UBound(vRngArr)
        Set oStdTerm = New CStdTerm
        oStdTerm.m_s������ = vRngArr(lRow, 1)
        oStdTerm.m_s�ܾ�������� = vRngArr(lRow, 2)
        oStdTerm.m_s������ = vRngArr(lRow, 3)
        oStdTerm.m_s���� = vRngArr(lRow, 4)
        oStdTerm.m_s�����γ��� = vRngArr(lRow, 5)
        oStdTerm.m_s������Ÿ�Ը� = vRngArr(lRow, 6)
        oStdTerm.m_i���� = CInt(vRngArr(lRow, 7))
        oStdTerm.m_i���� = CInt(vRngArr(lRow, 8))
        oStdTerm.m_s���Ǿ��� = vRngArr(lRow, 9)
        oStdTerm.m_s������Ÿ�Ա��̸� = vRngArr(lRow, 10)
'        oStdTerm.m_s������� = vRngArr(lRow, 8)
'        oStdTerm.m_s������ = vRngArr(lRow, 9)
'        oStdTerm.m_sǥ�شܾ����� = vRngArr(lRow, 10)
'        oStdTerm.m_s������Ÿ�Ա��̸� = vRngArr(lRow, 11)
        oStdTerm.SetData m_oStdWordDicP

On Error Resume Next
        m_oStdTermDic.Add oStdTerm.m_s������, oStdTerm
        If Err <> 0 Then '��� ���� �ߺ� �߻�
            Set oStdTermTmp = m_oStdTermDic.Item(oStdTerm.m_s������)
            sInfoMsg = "������ ������ ������ �ߴܵǾ����ϴ�: " & "������ �ߺ�" & vbLf & _
                       "�� �ߺ� �׸�" & vbLf & _
                           vbTab & "������: " & oStdTerm.m_s������ & vbLf & _
                           vbTab & "������: " & oStdTerm.m_s������ & vbLf & _
                       "�� ���� �׸�" & vbLf & _
                           vbTab & "������: " & oStdTermTmp.m_s������ & vbLf & _
                           vbTab & "������: " & oStdTermTmp.m_s������ & vbLf & vbLf & _
                       "�ߺ��� �����ϰ� ��� �����Ͻðڽ��ϱ�?"

            If IsOkToGo(sInfoMsg, "��� ���� �ߺ�") Then
                Resume Next
            Else
                End
                Exit For
            End If
        End If

        Err.Clear

        '2021-02-07 �߰�
        If m_o��������Dic.Exists(oStdTerm.m_sSorted�ܾ��������) Then
            '����� m_sSorted�ܾ������������ Item�� �����ϴ� ���
            Set oStdTermCol = m_o��������Dic.Item(oStdTerm.m_sSorted�ܾ��������)
        Else '�������� �ʴ� ���
            Set oStdTermCol = New Collection
            m_o��������Dic.Add oStdTerm.m_sSorted�ܾ��������, oStdTermCol
        End If
        oStdTermCol.Add oStdTerm
    Next lRow

    Dim v�������� As Variant, sChkResultTemp As String, sChkResult As String
    For Each v�������� In m_o��������Dic.Keys()
        Set oStdTermCol = m_o��������Dic.Item(v��������)
        sChkResultTemp = ""
        If oStdTermCol.Count > 1 Then
            lRow = 0
            Set oStdTerm = oStdTermCol.Item(1)
            sChkResultTemp = "������: " + oStdTerm.m_s������
            For Each oStdTerm In oStdTermCol
                lRow = lRow + 1
                sChkResultTemp = sChkResultTemp + vbLf + _
                                 "  - ������(" + CStr(lRow) + "): " + oStdTerm.m_s������
            Next oStdTerm
            sChkResult = sChkResult + sChkResultTemp + vbLf
        End If
    Next v��������
    m_s�����ߺ����˰�� = sChkResult
End Sub

Public Function Exists(aTerm As String) As Boolean
    Exists = m_oStdTermDic.Exists(aTerm)
End Function

Public Function Item(aTerm As String) As CStdTerm
    Set Item = m_oStdTermDic.Item(aTerm)
End Function

