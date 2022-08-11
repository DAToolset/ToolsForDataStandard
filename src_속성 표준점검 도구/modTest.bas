Attribute VB_Name = "modTest"
Option Explicit

Public Sub TestVariant1()
    Dim vOutRngArr As Variant
    vOutRngArr = Range("A1:B2").Value2
    vOutRngArr(1, 1) = "1,1"
    vOutRngArr(1, 2) = "1,2"
    vOutRngArr(2, 1) = "2,1"
    vOutRngArr(2, 2) = "2,2"

    Range("A1:B2").Value2 = vOutRngArr
End Sub

Public Sub TestVariant2()
    Dim vInRngArr As Variant, vOutRngArr As Variant, oOutRange As Range
    vInRngArr = Range("B4:C10")

    vInRngArr(1, 1) = "1,1"
    vInRngArr(1, 2) = "1,2"
    
    'ReDim vOutRngArr(LBound(vInRngArr, 1) To UBound(vInRngArr, 1), LBound(vInRngArr, 2) To UBound(vInRngArr, 2))

    Set oOutRange = Range("E4").Resize(2, 8) '쓰는 범위: 8개 컬럼
    'vOutRngArr = oOutRange.Value2
    ReDim vOutRngArr(oOutRange.Rows.Count - 1, oOutRange.Columns.Count - 1)
    'ReDim vOutRngArr(2, 8)
    
End Sub


Public Sub TestDupItemInCollection()
    Dim oCol As Collection
    Set oCol = New Collection
    oCol.Add "A"
    oCol.Add "A"
    Set oCol = Nothing
End Sub

'Dictionary의 value type이 instance와 collection이 모두 가능한지, 판별 가능한지 테스트
Public Sub TestDictionaryValue()
    Dim oDic As Dictionary, oCol As Collection, oWord1 As CStdWord, oWord2 As CStdWord, oObj1 As Object, oObj2 As Object
    Set oDic = New Dictionary: Set oCol = New Collection
    Set oWord1 = New CStdWord: Set oWord2 = New CStdWord
    oWord1.m_s단어논리명 = "사전": oWord2.m_s단어논리명 = "사전"
    oCol.Add oWord1
    oCol.Add oWord2
    oCol.Add oWord2
    oDic.Add "Instance", oWord1
    oDic.Add "Collection", oCol
    
    Set oObj1 = oDic("Instance")
    If TypeOf oObj1 Is CStdWord Then Debug.Print "Instance"
    Set oObj2 = oDic("Collection")
    If TypeOf oObj2 Is Collection Then Debug.Print "Collection"
    Debug.Print oObj2.Count
    Debug.Print oObj2(1).m_s단어논리명

    Set oDic = Nothing
    Set oCol = Nothing
    Set oWord1 = Nothing
    Set oWord2 = Nothing
    Set oObj1 = Nothing
    Set oObj2 = Nothing
End Sub

'Collection의 Key가 존재하지 않을 때 테스트
Public Sub TestCollection1()
    Dim oCol As Collection
    Set oCol = New Collection
    oCol.Add "A123", "A"
    oCol.Add "B123", "B"
    
    Debug.Print oCol("A")

    Debug.Print oCol("C")

    Set oCol = Nothing
End Sub

'Collection item의 값 변경 테스트
Public Sub TestCollection2()
    Dim oCol As Collection
    Set oCol = New Collection
    oCol.Add "A123", "A"

    oCol(1) = "X123" ' Error

    Set oCol = Nothing
End Sub

'String Array Item의 값 변경 테스트
Public Sub TestArrary1()
    Dim saMatchResult() As String

    ReDim Preserve saMatchResult(0 To 0)
    saMatchResult(0) = "A123"
    saMatchResult(0) = "X123"
    
    ReDim Preserve saMatchResult(0 To 1)
    saMatchResult(1) = "T123"
    
    ReDim Preserve saMatchResult(UBound(saMatchResult) + 1)
    ReDim saMatchResult(0)
    
End Sub

Function copyFilteredData() As Variant
    Dim selectedData() As Variant
    Dim aCnt As Long
    Dim rCnt As Long

    Range("B1").CurrentRegion.SpecialCells(xlCellTypeVisible).Select
    On Error GoTo MakeArray:
    For aCnt = 1 To Selection.Areas.Count
        For rCnt = 1 To Selection.Areas(aCnt).Rows.Count
            ReDim Preserve selectedData(UBound(selectedData) + 1)
            selectedData(UBound(selectedData)) = Selection.Areas(aCnt).Rows(rCnt)
        Next
    Next

    copyFilteredData = selectedData
    Exit Function

MakeArray:
    ReDim selectedData(1)
    Resume Next

End Function

Private Sub JoinArrayTest1()
    Dim sData As String, saData() As String
    sData = "ABC_DEF_GHI"
    Debug.Print sData
    saData = Split(sData)
    Debug.Print Join(saData, "_")
End Sub
