VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdWordDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ǥ�شܾ� Dictionary Class
Option Explicit

Public m_oStdWordDic As Dictionary 'Key: �ܾ����, Value: CStdWord Collection(2018-05-07) '������: CStdWord instance
Public m_oStdWordDicP As Dictionary 'Key: �ܾ����, Value: CStdWord Collection(2018-05-07) '������: CStdWord instance

Public m_eStdDicMatchOption As StdDicMatchOption

Private Sub Class_Initialize()
    Set m_oStdWordDic = New Dictionary
    Set m_oStdWordDicP = New Dictionary
End Sub

Private Sub Class_Terminate()
    m_oStdWordDic.RemoveAll
    m_oStdWordDicP.RemoveAll

    Set m_oStdWordDic = Nothing
    Set m_oStdWordDicP = Nothing
End Sub

'�ܾ���� Loading(By Variant Array)
Public Sub Load(aBaseRange As Range, aIsAllowDupWordLogicalName As Boolean, aIsAllowDupWordPhysicalName As Boolean)
On Error Resume Next
'    '�� ã�� ��� �ܾ���� Loading���� ����
'    If m_eStdDicMatchOption = TermOnly Then Exit Sub

    '��Ͽ� �ƹ� ���� ���� ��� exit
    If Trim(aBaseRange.Offset(1, 0)) = "" Then Exit Sub

    Dim vRngArr As Variant
    vRngArr = Range(aBaseRange, aBaseRange.End(xlDown)).Resize(, 9).Value2 '�д� ����: 6�� �÷�
    Dim lRow As Long, oStdWord As CStdWord, oStdWordTmp As CStdWord, sInfoMsg As String
    Dim bIs�ܾ�������� As Boolean, oStdWordCol As CStdWordCol, oExistedStdObj As Object

    For lRow = LBound(vRngArr) To UBound(vRngArr)
        bIs�ܾ�������� = False
        Set oStdWord = New CStdWord
        oStdWord.SetStdWordDic Me
        oStdWord.m_s�ܾ���� = vRngArr(lRow, 1)
        oStdWord.m_s�ܾ���� = vRngArr(lRow, 2)
        oStdWord.m_s�ܾ���� = vRngArr(lRow, 3)
        oStdWord.m_s�ܾ�� = vRngArr(lRow, 4)
        'oStdWord.m_s�ܾ����� = vRngArr(lRow, 5)
        oStdWord.m_bǥ�ؿ��� = GetBoolean(vRngArr(lRow, 5))
        oStdWord.m_b�Ӽ��з���� = GetBoolean(vRngArr(lRow, 6))
        oStdWord.m_sǥ�س��� = vRngArr(lRow, 7)
        oStdWord.m_s���Ǿ� = vRngArr(lRow, 8)
        oStdWord.m_s�����κз��� = vRngArr(lRow, 9)

        '�ܾ���� ���� Dictionary Build
        bIs�ܾ�������� = m_oStdWordDic.Exists(oStdWord.m_s�ܾ����)
        If bIs�ܾ�������� Then
            If aIsAllowDupWordLogicalName Then
                '�ܾ� ���� �ߺ�(�������Ǿ�) ���
                If (oStdWord.m_bǥ�ؿ��� = False) And _
                   (oStdWord.m_s�ܾ���� <> oStdWord.m_sǥ�س���) Then
                    GoTo Skip_�ܾ����Build 'ǥ�ؿ��� False�̰� ����� ǥ�س����� �������� ���� ��� Skip
                End If

                oStdWord.m_b�����ߺ����� = True
                '�޽��� �������� �ʰ� Collection���� ���� Value item ��ü
                Set oExistedStdObj = m_oStdWordDic.Item(oStdWord.m_s�ܾ����) '���� �ܾ�
                If TypeOf oExistedStdObj Is CStdWord Then
                    '���� Value�� CStdWord instance�� ��� --> CStdWordCol �����ϰ� ��ü
                    Set oStdWordCol = New CStdWordCol
                    oStdWordCol.Add oExistedStdObj 'Collection�� ���� �ܾ��߰�
                    oStdWordCol.Add oStdWord 'Collection�� ���� �ܾ� �߰�

                    m_oStdWordDic.Remove oStdWord.m_s�ܾ���� '���� Key-Value(Instance) ����
                    m_oStdWordDic.Add oStdWord.m_s�ܾ����, oStdWordCol '���� Key-Value(Collection) �߰�

                ElseIf TypeOf oExistedStdObj Is CStdWordCol Then
                    '���� Value�� CStdWord Collection�� ��� --> �̹� �������Ǿ �ִ� ���. ���� �ܾ CStdWordCol�� Add
                    oExistedStdObj.Add oStdWord
                End If
            Else
                '�ܾ� ���� �ߺ�(�������Ǿ�) �����
                Set oStdWordTmp = m_oStdWordDic.Item(oStdWord.m_s�ܾ����)
                sInfoMsg = "������ ������ ������ �ߴܵǾ����ϴ�: " & "�ܾ���� �ߺ�" & vbLf & _
                           "�� �ߺ� �׸�" & vbLf & _
                               vbTab & "�ܾ����: " & oStdWord.m_s�ܾ���� & vbLf & _
                               vbTab & "�ܾ����: " & oStdWord.m_s�ܾ���� & vbLf & _
                           "�� ���� �׸�" & vbLf & _
                               vbTab & "�ܾ����: " & oStdWordTmp.m_s�ܾ���� & vbLf & _
                               vbTab & "�ܾ����: " & oStdWordTmp.m_s�ܾ����
                If IsOkToGo(sInfoMsg, "�ܾ� ���� �ߺ�") Then
                    'Resume Next
                Else
                    End
                     Exit For
                End If
            End If
        Else
            If (oStdWord.m_bǥ�ؿ��� = False) And _
               (oStdWord.m_sǥ�س��� = "") Then
                GoTo Skip_�ܾ����Build 'ǥ�ؿ��� False�̰� ǥ�س����� ���� ��� Skip
            End If
            '�ܾ���� ���� ���
            m_oStdWordDic.Add oStdWord.m_s�ܾ����, oStdWord
        End If

Skip_�ܾ����Build:

On Error Resume Next
        '�ܾ���� ���� Dictionary Build
        If oStdWord.m_bǥ�ؿ��� = True Then 'ǥ�شܾ��ϰ�츸 ������ ���� Dictionary�� �߰�(������ �ߺ� ����)
            If m_oStdWordDicP.Exists(oStdWord.m_s�ܾ����) Then
                oStdWord.m_b�������ߺ����� = True '�ߺ��� �������� �ִ��� üũ
                m_oStdWordDicP(oStdWord.m_s�ܾ����).m_b�������ߺ����� = True
            End If
            m_oStdWordDicP.Add oStdWord.m_s�ܾ����, oStdWord
            If Err <> 0 And aIsAllowDupWordPhysicalName = False Then '�ܾ� ������ �ߺ� �߻�
                Set oStdWordTmp = m_oStdWordDicP.Item(oStdWord.m_s�ܾ����)
                sInfoMsg = "������ ������ ������ �ߴܵǾ����ϴ�: " & "�ܾ���� �ߺ�: " & vbLf & _
                           "�� �ߺ� �׸�" & vbLf & _
                               vbTab & "�ܾ����: " & oStdWord.m_s�ܾ���� & vbLf & _
                               vbTab & "�ܾ����: " & oStdWord.m_s�ܾ���� & vbLf & _
                           "�� ���� �׸�" & vbLf & _
                               vbTab & "�ܾ����: " & oStdWordTmp.m_s�ܾ���� & vbLf & _
                               vbTab & "�ܾ����: " & oStdWordTmp.m_s�ܾ���� & vbLf & vbLf & _
                           "�ߺ��� �����ϰ� ��� �����Ͻðڽ��ϱ�?"
    
                If IsOkToGo(sInfoMsg, "�ܾ� ������ �ߺ�") Then
                    Resume Next
                Else
                    End
                    Exit For
                End If
            End If
        End If
On Error GoTo 0

    Next lRow

On Error GoTo 0
End Sub

'���� ������ Item ã��
Public Function Exists(aWord As String) As Boolean
    Exists = m_oStdWordDic.Exists(aWord)
End Function

'������ ������ Item ã��
Public Function ExistsP(aWord As String) As Boolean
    ExistsP = m_oStdWordDicP.Exists(aWord)
End Function

Public Function Item(aWord As String) As Object
    Set Item = m_oStdWordDic.Item(aWord)
End Function

Public Function ItemP(aWord As String) As CStdWord
    Set ItemP = m_oStdWordDicP.Item(aWord)
End Function

Public Function Count() As Long
    Count = m_oStdWordDic.Count
End Function
