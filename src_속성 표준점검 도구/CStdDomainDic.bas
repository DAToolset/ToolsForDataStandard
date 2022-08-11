VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdDomainDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_oStdDomainDic As Dictionary 'Key: �����κз���, Value: CStdDomain�� Collection
Public m_oStdDomainDicT As Dictionary 'Key: �����κз���, Value: Dictionary(Key:m_s������Ÿ�Ա��̸�, Value: CStdDomain�� Collection)
Public m_oStdWordDic As CStdWordDic '�ܾ����(�Ӽ��з����� �����κз����� Ȯ���ϴ� �뵵)
Public m_eStdDicMatchOption As StdDicMatchOption

Private Sub Class_Initialize()
    Set m_oStdDomainDic = New Dictionary
    Set m_oStdDomainDicT = New Dictionary
End Sub

Private Sub Class_Terminate()
    m_oStdDomainDic.RemoveAll
    Set m_oStdDomainDic = Nothing

    m_oStdDomainDicT.RemoveAll
    Set m_oStdDomainDicT = Nothing
End Sub

Public Sub SetStdWordDic(aStdWordDic As CStdWordDic)
    Set m_oStdWordDic = aStdWordDic
End Sub

Public Sub Add(aStdDomain As CStdDomain)
    Dim oStdDomainCol As Collection
    If m_oStdDomainDic.Exists(aStdDomain.m_s�����κз���) Then
        Set oStdDomainCol = m_oStdDomainDic(aStdDomain.m_s�����κз���)
    Else
        Set oStdDomainCol = New Collection
        m_oStdDomainDic.Add aStdDomain.m_s�����κз���, oStdDomainCol
    End If
    oStdDomainCol.Add aStdDomain
End Sub

Public Function GetDomainCollection(a�����κз��� As String) As Collection
    If m_oStdDomainDic.Exists(a�����κз���) Then
        Set GetDomainCollection = m_oStdDomainDic(a�����κз���)
    End If
End Function

Public Sub Load(aBaseRange As Range)
    Dim oStdDomain As CStdDomain
    Dim lRow As Long

    '��Ͽ� �ƹ� ���� ���� ��� exit
    If Trim(aBaseRange.Offset(1, 0)) = "" Then Exit Sub

    Dim vRngArr As Variant
    vRngArr = Range(aBaseRange, aBaseRange.End(xlDown)).Resize(, 8).Value2 '�д� ����: 8�� �÷�

    For lRow = LBound(vRngArr) To UBound(vRngArr)
        Set oStdDomain = New CStdDomain
        With oStdDomain
            .m_s�����κз��� = vRngArr(lRow, 1)
            .m_s�����γ��� = vRngArr(lRow, 2)
            .m_s�����ι����� = vRngArr(lRow, 3)
            .m_s�����μ��� = vRngArr(lRow, 4)
            .m_s������Ÿ�Ը� = vRngArr(lRow, 5)
            .m_i���� = vRngArr(lRow, 6)
            .m_i���� = vRngArr(lRow, 7)
            .m_s������Ÿ�Ա��̸� = vRngArr(lRow, 8)
            '.m_s������Ÿ�Ա��̸� = GetDataTypeStr(.m_s������Ÿ�Ը�, .m_i����, .m_i����)
        End With

        Me.Add oStdDomain
    Next lRow

    '������Ÿ�Ա��̸� �˻��� Dictionary build
    Dim oKey As Variant, s�����κз��� As String
    Dim oStdDomainCollection As Collection, oTypeDic As Dictionary
    For Each oKey In m_oStdDomainDic.Keys
        s�����κз��� = oKey
        'Set oStdDomain = m_oStdDomainDic(s�����κз���)
        Set oStdDomainCollection = m_oStdDomainDic(s�����κз���)
        For Each oStdDomain In oStdDomainCollection
            If m_oStdDomainDicT.Exists(s�����κз���) Then
                Set oTypeDic = m_oStdDomainDicT(s�����κз���)
            Else
                Set oTypeDic = New Dictionary
                m_oStdDomainDicT.Add s�����κз���, oTypeDic
            End If
On Error Resume Next '�����κз��� ���� ������Ÿ�Ա��̸��� �ߺ��ǵ� ����
            oTypeDic.Add oStdDomain.m_s������Ÿ�Ա��̸�, oStdDomain
On Error GoTo 0
        Next
    Next
End Sub

'�Ӽ�������Ÿ�Ա��̸�� ǥ�ؿ���� ������Ÿ�Ա��̸��� ���� ��� return
'Public Function GetCheckAttrDataType(a�Ӽ������ As String, _
'       a�Ӽ�������Ÿ�Ա��̸� As String, aǥ�ؿ�����Ÿ�Ա��̸� As String)
Public Function GetCheckAttrDataType(a�Ӽ������ As String, a�Ӽ������StdWord As CStdWord, _
        a�Ӽ�������Ÿ�Ա��̸� As String, aǥ�ؿ�����Ÿ�Ա��̸� As String, _
        Optional aStdWord As CStdWord) As String

    Dim sResult As String, oStdDomain As CStdDomain, oStdDomainTerm As CStdDomain, oStdDomainAtt As CStdDomain
    Dim s�Ӽ������_�����κз��� As String, oStdWord As CStdWord, b�Ӽ������_�����κз����翩�� As Boolean

    If a�Ӽ�������Ÿ�Ա��̸� = "" Then
        sResult = "�Ӽ� Data Type ���� �ʿ�"
        GetCheckAttrDataType = sResult
        Exit Function
    End If

    Set oStdDomainAtt = New CStdDomain '�Ӽ� ������Ÿ��(�񱳿�)
    oStdDomainAtt.SetDomain a�Ӽ�������Ÿ�Ա��̸�

    If aǥ�ؿ�����Ÿ�Ա��̸� > "" Then
        '------------------------------------------------------------------------------------------
        'ǥ�ؿ��� ��Ī�� ���
        Set oStdDomainTerm = New CStdDomain 'ǥ�ؿ�����Ÿ��(�񱳿�)
        oStdDomainTerm.SetDomain aǥ�ؿ�����Ÿ�Ա��̸�

        sResult = GetCompareResult("ǥ�ؿ��", oStdDomainAtt, oStdDomainTerm)
    Else
        '------------------------------------------------------------------------------------------
        'ǥ�شܾ�� ��Ī�� ���

        sResult = "������ ���� ���"
        '�Ӽ������� �����κз� ã��
        s�Ӽ������_�����κз��� = a�Ӽ������
        If Not a�Ӽ������StdWord Is Nothing Then
            s�Ӽ������_�����κз��� = a�Ӽ������StdWord.m_s�����κз���
        Else
            sResult = sResult + vbLf + "�Ӽ������: ǥ�شܾ� ����(" & a�Ӽ������ & ")"
        End If

        Set oStdDomain = GetStdDomain(a�����κз���:=s�Ӽ������_�����κз��� _
                                , a������Ÿ�Ա��̸�:=a�Ӽ�������Ÿ�Ա��̸� _
                                , a�����κз������翩��:=b�Ӽ������_�����κз����翩��)

        If b�Ӽ������_�����κз����翩�� = False Then
            sResult = sResult + vbLf + "�Ӽ������:�����κз� ����"
        End If

        If Not oStdDomain Is Nothing Then
            '�Ӽ� �з�� ������ ������Ÿ�� �� �Ӽ� ������ Ÿ���� ������
            sResult = GetCompareResult("������", oStdDomainAtt, oStdDomain)
        Else
            sResult = sResult + vbLf + "������ �߰� �ʿ� <" + _
                      IIf(s�Ӽ������_�����κз��� > "", s�Ӽ������_�����κз���, a�Ӽ������) + _
                      ">: " + a�Ӽ�������Ÿ�Ա��̸�
        End If
    End If
    '----------------------------------------------------------------------------------------------

    GetCheckAttrDataType = sResult
End Function

Public Function GetStdDomain(a�����κз��� As String, _
        a������Ÿ�Ա��̸� As String, a�����κз������翩�� As Boolean) As CStdDomain
    Dim oResult As CStdDomain, oTypeDic As Dictionary
    a�����κз������翩�� = False
    If m_oStdDomainDicT.Exists(a�����κз���) Then
        a�����κз������翩�� = True
        Set oTypeDic = m_oStdDomainDicT(a�����κз���)
        If oTypeDic.Exists(a������Ÿ�Ա��̸�) Then Set oResult = oTypeDic(a������Ÿ�Ա��̸�)
    End If
    Set GetStdDomain = oResult
End Function

Public Function ExistsDomainDataType(a�����κз��� As String, _
        a������Ÿ�Ա��̸� As String) As Boolean
    Dim bResult As Boolean, oTypeDic As Dictionary
    bResult = False
    If m_oStdDomainDicT.Exists(a�����κз���) Then
        Set oTypeDic = m_oStdDomainDicT(a�����κз���)
        If oTypeDic.Exists(a������Ÿ�Ա��̸�) Then bResult = True
    End If
    ExistsDomainDataType = bResult
End Function

'�� Domain�� TypeSize �񱳰�� return
'aDomainAtt: �񱳱��� Domain (�Ӽ����� Type/Size)
'aDomainTgt: �񱳴�� Domain (��� �Ǵ� �Ӽ��з��� Type/Size)
Public Function GetCompareResult(aCompareType As String, _
            aDomainAtt As CStdDomain, aDomainTgt As CStdDomain) As String
    Dim sResult As String
    sResult = aCompareType + " Type/Size �� ���"
    If aDomainAtt.m_s������Ÿ�Ը� <> aDomainTgt.m_s������Ÿ�Ը� Then
        sResult = sResult + vbLf + "Ÿ�� ����ġ"
    End If
    If aDomainAtt.m_i���� <> aDomainTgt.m_i���� Then
        sResult = sResult + vbLf + "���� ����ġ"
        If aDomainAtt.m_i���� > aDomainTgt.m_i���� Then '�Ӽ��� Size�� �� ū ���(������ �߰� �Ǵ� �Ӽ� size ����)
            sResult = sResult + "(����! ������ �߰� �Ǵ� �Ӽ� Size ���� �ʿ�)"
        ElseIf aDomainAtt.m_i���� < aDomainTgt.m_i���� Then '�� ��� Domain Size�� �� ū ���(��κ��� ��������)
            sResult = sResult + "(���� Ȯ��)"
        End If
    ElseIf aDomainAtt.m_i���� <> aDomainTgt.m_i���� Then
        sResult = sResult + vbLf + "�Ҽ��� ���� ����ġ"
        If aDomainAtt.m_i���� > aDomainTgt.m_i���� Then '�Ӽ��� Size�� �� ū ���(������ �߰� �Ǵ� �Ӽ� size ����)
            sResult = sResult + "(����! ������ �߰� �Ǵ� �Ӽ� Size ���� �ʿ�)"
        ElseIf aDomainAtt.m_i���� < aDomainTgt.m_i���� Then '�� ��� Domain Size�� �� ū ���(��κ��� ��������)
            sResult = sResult + "(���� Ȯ��)"
  End If
    Else
  sResult = aCompareType + " Type/Size ��ġ"
    End If
    GetCompareResult = sResult
End Function


''�� Domain�� TypeSize �񱳰�� return
''aDomainAtt: �񱳱��� Domain (�Ӽ����� Type/Size)
''aDomainTgt: �񱳴�� Domain (��� �Ǵ� �Ӽ��з��� Type/Size)
'Public Function GetCompareResult(aCompareType As String, _
'            aDomainAtt As CStdDomain, aDomainTgt As CStdDomain) As String
'    Dim sResult As String
'    sResult = aCompareType + " Type/Size �� ���"
'    If aDomainAtt.m_s������Ÿ�Ը� <> aDomainTgt.m_s������Ÿ�Ը� Then
'        sResult = sResult + vbLf + "Ÿ�� ����ġ"
'    ElseIf aDomainAtt.m_i���� <> aDomainTgt.m_i���� Then
'        sResult = sResult + vbLf + "���� ����ġ"
'        If aDomainAtt.m_i���� > aDomainTgt.m_i���� Then '�Ӽ��� Size�� �� ū ���(������ �߰� �Ǵ� �Ӽ� size ����)
'            sResult = sResult + "(������ �߰� �Ǵ� �Ӽ� Size ���� �ʿ�)"
''        Else '�� ��� Domain Size�� �� ū ���(��κ��� ��������)
''            sResult = sResult + ""
'        End If
'    ElseIf aDomainAtt.m_i���� <> aDomainTgt.m_i���� Then
'        sResult = sResult + vbLf + "�Ҽ��� ���� ����ġ"
'        If aDomainAtt.m_i���� > aDomainTgt.m_i���� Then '�Ӽ��� Size�� �� ū ���(������ �߰� �Ǵ� �Ӽ� size ����)
'            sResult = sResult + "(������ �߰� �Ǵ� �Ӽ� Size ���� �ʿ�)"
''        Else '�� ��� Domain Size�� �� ū ���(��κ��� ��������)
''            sResult = sResult + ""
'  End If
'    Else
'  sResult = aCompareType + " Type/Size ��ġ"
'    End If
'    GetCompareResult = sResult
'End Function
