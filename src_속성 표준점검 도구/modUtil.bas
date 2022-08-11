Attribute VB_Name = "modUtil"
Option Explicit
Private Declare PtrSafe Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare PtrSafe Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer

Const VK_SHIFT As Integer = &H10

'Split & Trim ó��
Public Function SplitTrim(aExpression As String, aDelimeter As String) As String()
    Dim saOut() As String, i As Integer
    saOut = Split(aExpression, aDelimeter)
    For i = 1 To UBound(saOut)
        saOut(i) = Trim(saOut(i))
    Next i

    SplitTrim = saOut
End Function

'Split & Trim �� n��° Item Return
'n�� ������ ��� �ڿ������� (���)n��° Item Return
Public Function SplitAndGetNItem(aText As String, aDelimeter As String, aNth As Integer) As String
    If aText = "" Then Exit Function
    Dim saToken() As String
    saToken = SplitTrim(aText, aDelimeter)
    If aNth < 0 Then
        SplitAndGetNItem = saToken(aNth + UBound(saToken) + 1)
    Else
        SplitAndGetNItem = saToken(aNth)
    End If
End Function

'Text�� �޺κ� ���� return
Public Function GetNumberSuffix(aText As String) As String
    Dim sResult As String, sToken As String, iLen As Integer, iTextLen As Integer
    iTextLen = Len(aText)
    sResult = ""
    For iLen = iTextLen To 1 Step -1
        sToken = Mid(aText, iLen, 1)
        Select Case sToken
            Case "0" To "9"
                sResult = sToken + sResult
            Case Else
                Exit For
        End Select
    Next iLen
    GetNumberSuffix = sResult
End Function

'Text�� suffix�� �����ϰ� return
Public Function GetTextWithoutSuffix(aText As String, aSuffix As String) As String
    Dim iSuffixIdx As Integer
    iSuffixIdx = InStrRev(aText, aSuffix)
    If iSuffixIdx > 0 Then
        GetTextWithoutSuffix = Mid(aText, 1, iSuffixIdx - 1)
    End If
End Function

'DataType Name, Precision, Scale�� ������ ���ڿ� return
'Public Function GetDataTypeStr(aDataType As String, aPrecision As String, aScale As String) As String
Public Function GetDataTypeStr(aDataType As String, aPrecision As Integer, aScale As Integer) As String
    Dim sResult As String
    sResult = IIf(aDataType = "VARCHAR", "VARCHAR2", aDataType)
    sResult = sResult + IIf(aPrecision > 0, "(" + CStr(aPrecision), "") 'Precision
    sResult = sResult + IIf(aScale > 0, "," + CStr(aScale), "") 'aScale
    sResult = sResult + IIf(aPrecision > 0, ")", "")
    GetDataTypeStr = sResult
End Function

'OutputDebugString API�� �̿��� Debug Message ���
'DebugView���� �̿��Ͽ� �޽��� View ������
Public Sub DoLog(aMsg As String)
    OutputDebugString "[STD]" & aMsg
End Sub

'Query ������ Sheet�� ���ļ���
Public Sub DoResultFormatting(aSht As Worksheet)
    aSht.Range("A1").Activate
    aSht.Cells.Select
    With Selection.Font
        .Name = "���� ���"
        .Size = 9
    End With

    aSht.Cells.EntireColumn.AutoFit
    ActiveWindow.DisplayGridlines = False
    'Range("A1:E47").Select
    ActiveCell.CurrentRegion.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    aSht.Range("A1").Select
    aSht.Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    aSht.Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub

Public Sub CheckNDelete(aBook As Workbook, aSheetName As String)
On Error GoTo NoSuchSheet

    Application.DisplayAlerts = False

    If Len(aBook.Sheets(aSheetName).Name) > 0 Then
        aBook.Sheets(aSheetName).Delete
    End If

NoSuchSheet:

    Application.DisplayAlerts = True
End Sub

'Worksheet�� SheetName�� �����ϴ��� Ȯ��
Public Function IsSheetExists(aSheetName As String) As Boolean
    Dim oSht As Worksheet

    On Error Resume Next
    Set oSht = ThisWorkbook.Sheets(aSheetName)
    On Error GoTo 0
    IsSheetExists = Not oSht Is Nothing
End Function

'Control size�� ���� ����(��ư�� Ŭ���� ������ ũ�Ⱑ Ŀ���� ���� ���� �ذ�å)
Public Sub ResetControlSize(aWorkSheet As Worksheet)
     ' Kluge to reset button and font size to resolve shrinking/enlarging buttons
     ' Bisham Singh 2011 12 23
     ' This is sledge hammmer approach. If there is a better one, please let us know.
     ' I run this everytime a button is triggered.
    Dim Shape, mWorkSheet As Worksheet, mOLE As OLEObject
    Application.ScreenUpdating = False
    Set mWorkSheet = aWorkSheet
    For Each mOLE In mWorkSheet.OLEObjects
        If mOLE.Name = "chkRefreshStdDic" Then GoTo Skip_Control
        If (TypeName(mOLE.Object) = "Label") And (mOLE.Object.Caption = " ") Then GoTo Skip_Control
        mOLE.Width = mOLE.Width
        mOLE.Height = mOLE.Height
        mOLE.Object.FontSize = mOLE.Object.FontSize
        'mOLE.Object.AutoSize = mOLE.Object.AutoSize
        mOLE.Object.AutoSize = False
        mOLE.Object.AutoSize = True
Skip_Control:
    Next
    Application.ScreenUpdating = True
End Sub

'Ȯ�� �޽��� �����ְ� Yes Ŭ���� True, �׿ܴ� False
Public Function IsOkToGo(aMsg As String, Optional aTitle As String = "") As Boolean
    IsOkToGo = False
    Dim iMsgResult As VbMsgBoxResult, sTitle As String
    sTitle = IIf(aTitle = "", "Ȯ��", aTitle)
    iMsgResult = MsgBox(aMsg, vbYesNo + vbQuestion + vbDefaultButton1, sTitle)
    If iMsgResult = vbYes Then IsOkToGo = True
End Function

'��� ����
Public Sub DoClearList(aRange As Range, Optional aIsForce As Boolean = False)
    If Not aIsForce Then
        Dim iMsgResult As VbMsgBoxResult
        iMsgResult = MsgBox("����� �ʱ�ȭ�մϴ�." & vbLf & "��� �����Ͻðڽ��ϱ�?", _
                vbYesNo + vbQuestion + vbDefaultButton1, "Ȯ��")
        If iMsgResult <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    aRange.Worksheet.Activate
    aRange.Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    aRange.Select
    Application.ScreenUpdating = True
End Sub

'Y/N ���� ������ boolean �Ǻ�
Public Function GetBoolean(aVal As Variant) As Boolean
    Select Case aVal
        Case "Y", "y"
            GetBoolean = True
        Case Else
            GetBoolean = False
    End Select
End Function

'��ũ�� ���ϸ��� Version String ��������
Public Function GetVersionString() As String
    Dim sFileName As String, sVersionString As String, sChar As String
    Dim lIdx As Long, lFIdx As Long, lTIdx As Long, lLen As Long
    sFileName = ThisWorkbook.Name '��: �Ӽ��� ǥ������ ����_v1.21_20181209_1.xlsm

    lFIdx = InStr(1, sFileName, "V", vbTextCompare)
    If lFIdx <= 0 Then
        sVersionString = C_VERSION_STRING
    Else
        For lTIdx = lFIdx To Len(sFileName)
            sChar = Mid(sFileName, lTIdx, 1)
            If (IsNumeric(sChar)) Or _
               (UCase(sChar) = "V") Or _
               (sChar = ".") Then
                'Version String�� �����Ǵ� ������ ���
            Else
                Exit For
            End If
        Next lTIdx
        lLen = lTIdx - lFIdx
        sVersionString = Mid(sFileName, lFIdx, lLen)
    End If
    GetVersionString = sVersionString
End Function

'Shift Key�� ���ȴ��� üũ
Public Function IsShiftKeyDown() As Boolean
    If GetKeyState(VK_SHIFT) < 0 Then IsShiftKeyDown = True Else IsShiftKeyDown = False
End Function

'QuickSort
'Example: Call QuickSort(myArray, 0, UBound(myArray))
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub
