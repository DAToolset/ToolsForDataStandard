Option Explicit
Dim bIsWantToStop As Boolean

Private Sub cmdRunDicSearch_Click()
    Range("A1").Select
    DoEvents
    
    Dim bIsKorDicSearch As Boolean, bIsEngDicSearch As Boolean, sTargetDic As String
    bIsKorDicSearch = chkKorDic.Value: bIsEngDicSearch = chkEngDic.Value
    If (Not bIsKorDicSearch) And (Not bIsEngDicSearch) Then
        MsgBox "검색 대상 사전중 적어도 1개는 선택해야 합니다", vbExclamation + vbOKOnly, "검색 대상 사전 확인"
        Exit Sub
    End If

    Dim bIsMatchTypeExact As Boolean, bIsMatchTypeTermOr As Boolean, bIsMatchTypeAllTerm As Boolean '검색결과 표시 설정
    bIsMatchTypeExact = chkMatchTypeExact.Value: bIsMatchTypeTermOr = chkMatchTypeTermOr.Value: bIsMatchTypeAllTerm = chkMatchTypeAllTerm.Value

    If (bIsMatchTypeExact Or bIsMatchTypeTermOr Or bIsMatchTypeAllTerm) = False Then
        MsgBox "검색결과 표시 설정중 적어도 하나는 선택해야 합니다.", vbExclamation + vbOKOnly, "확인"
        Exit Sub
    End If

    If bIsKorDicSearch And Not bIsEngDicSearch Then sTargetDic = "국어사전"
    If Not bIsKorDicSearch And bIsEngDicSearch Then sTargetDic = "영어사전"
    If bIsKorDicSearch And bIsEngDicSearch Then sTargetDic = "국어사전, 영어사전"
    
    Dim lMaxResultCount As Long
    lMaxResultCount = CInt(txtMaxResultCount.Value)

    If MsgBox("사전 검색을 시작하시겠습니까?" + vbLf + _
              "대상 사전: " + sTargetDic + vbLf + _
              "결과출력 제한개수: " + CStr(lMaxResultCount) _
              , vbQuestion + vbYesNoCancel, "확인") <> vbYes Then Exit Sub

    Dim i As Long, iResultOffset As Long
    bIsWantToStop = False
    DoEvents

    Dim sWord As String, oKorDicSearchResult As TDicSearchResult, oEngDicSearchResult As TDicSearchResult
    Dim oBaseRange As Range
    Set oBaseRange = Range("검색결과Header").Offset(1, 0)
    oBaseRange.Select
    For i = 0 To 100000
        If bIsWantToStop Then
            MsgBox "사용자의 요청으로 검색을 중단합니다.", vbInformation + vbOKOnly, "확인"
            Exit For
        End If
        If chkSkipIfResultExists.Value = True And _
           oBaseRange.Offset(i, 1) <> "" Then GoTo Continue_For '이미 내용이 있으면 Skip
        sWord = oBaseRange.Offset(i)
        If sWord = "" Then Exit For
        oBaseRange.Offset(i).Select

        Application.ScreenUpdating = False
        If bIsKorDicSearch Then '국어사전 검색결과 표시
            oKorDicSearchResult = DoDicSearch(dtsKorean, sWord, bIsMatchTypeExact, bIsMatchTypeTermOr, bIsMatchTypeAllTerm, lMaxResultCount)
            oBaseRange.Offset(i, 1).Select
            With oKorDicSearchResult
                oBaseRange.Offset(i, 1) = .sMatchType
                oBaseRange.Offset(i, 2) = .sSearchEntry
                oBaseRange.Offset(i, 3) = .sMeaning
                If oKorDicSearchResult.sLinkURL <> "" Then
                    With ActiveSheet.Hyperlinks.Add(Anchor:=oBaseRange.Offset(i, 4), Address:=.sLinkURL, TextToDisplay:="네이버국어사전 열기: " & .sLinkWord)
                        .Range.Font.Size = 8
                    End With
                End If
                oBaseRange.Offset(i, 5) = .sSynonymList
                oBaseRange.Offset(i, 6) = .sAntonymList
            
            End With
        End If

        If bIsEngDicSearch Then '영어사전 검색결과 표시
            oEngDicSearchResult = DoDicSearch(dtsEnglish, sWord, bIsMatchTypeExact, bIsMatchTypeTermOr, bIsMatchTypeAllTerm, lMaxResultCount)
            'oBaseRange.Offset(i, 7).Select
            With oEngDicSearchResult
                oBaseRange.Offset(i, 7) = .sMatchType
                oBaseRange.Offset(i, 8) = .sSearchEntry
                oBaseRange.Offset(i, 9) = .sMeaning
                If oKorDicSearchResult.sLinkURL <> "" Then
                    With ActiveSheet.Hyperlinks.Add(Anchor:=oBaseRange.Offset(i, 10), Address:=.sLinkURL, TextToDisplay:="네이버영어사전 열기: " & .sLinkWord)
                        .Range.Font.Size = 8
                    End With
                End If
                oBaseRange.Offset(i, 11) = .sSynonymList
                oBaseRange.Offset(i, 12) = .sAntonymList
            
            End With
        End If
        Application.ScreenUpdating = True

Continue_For:
        DoEvents
    Next i

    MsgBox "사전 검색을 완료하였습니다", vbOKOnly + vbInformation
End Sub


Private Sub Test()
    ' Advanced example: Read .json file and load into sheet (Windows-only)
    ' (add reference to Microsoft Scripting Runtime)
    ' {"values":[{"a":1,"b":2,"c": 3},...]}
    
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    Dim JsonText As String
    Dim Parsed As Dictionary

    ' Read .json file
    Set JsonTS = FSO.OpenTextFile("D:\Project\My\2014_엑셀VBA\MyVBA\네이버 사전 검색\네이버 검색결과_최초_ANSI.json", ForReading)
    JsonText = JsonTS.ReadAll
    JsonTS.Close

    ' Parse json to Dictionary
    ' "values" is parsed as Collection
    ' each item in "values" is parsed as Dictionary
    Set Parsed = JsonConverter.ParseJson(JsonText)
    
    ' Prepare and write values to sheet
'    Dim Values As Variant
'    ReDim Values(Parsed("searchResultMap").Count, 3)

    Dim oItem As Dictionary, oMeansCollector As Dictionary, oMeans As Dictionary
    Dim oSimWords As Dictionary, oAntWord As Dictionary
    Dim i As Long
    
    i = 0
    For Each oItem In Parsed("searchResultMap")("searchResultListMap")("WORD")("items")
        If oItem("matchType") <> "exact:entry" Then Exit For
        'Debug.Print oItem("matchType")
        i = i + 1
        
        Debug.Print "사전명: " & oItem("sourceDictnameKO")
        Debug.Print "사전 URL: " & oItem("destinationLink")

        For Each oMeansCollector In oItem("meansCollector")
            Debug.Print "품사: " & oMeansCollector("partOfSpeech")
            For Each oMeans In oMeansCollector("means")
                Debug.Print "뜻: " & oMeans("value")
            Next oMeans
        Next oMeansCollector

        For Each oSimWords In oItem("similarWordList")
            Debug.Print "유의어: " & oSimWords("similarWordName")
            Debug.Print "유의어 Link: " & oSimWords("similarWordLink")
        Next oSimWords

        For Each oAntWord In oItem("antonymWordList")
            Debug.Print "반의어: " & oAntWord("antonymWordName")
            Debug.Print "반의어 Link: " & oAntWord("antonymWordLink")
        Next oAntWord
    
    Next oItem
    'Debug.Print i
End Sub

Private Sub cmdClearSearchResult_Click()
    If MsgBox("사전검색 결과를 초기화합니다." + vbLf + "계속 진행하시겠습니까?", vbQuestion + vbYesNoCancel, "확인") = vbYes Then
        ClearResult
    End If
End Sub

Private Sub cmdHelp_Click()
    frmHelp.Show
End Sub

Private Sub cmdStopSearch_Click()
    If MsgBox("사전 검색을 중지하시겠습니까?", vbQuestion + vbYesNoCancel, "확인") = vbYes Then
        bIsWantToStop = True
    End If
End Sub

Private Sub ClearResult()
    Dim oCurRange As Range: Set oCurRange = ActiveCell
    Dim oBaseRange As Range: Set oBaseRange = Range("검색결과Header")
    Dim lColOffset As Long ': lColOffset = 5 'Clear할 컬럼의 Offset(= 갯수-1)
    Dim lRowOffset As Long
    Application.ScreenUpdating = False
    oBaseRange.Select
    Range(oBaseRange, oBaseRange.End(xlDown)).Select
    lRowOffset = Selection.Rows.Count - 1
    Range(oBaseRange, oBaseRange.End(xlToRight)).Select
    lColOffset = Selection.Columns.Count - 1

    Set oBaseRange = oBaseRange.Offset(1, 1)
    Range(oBaseRange, oBaseRange.Offset(lRowOffset, lColOffset)).ClearContents
    'Selection.ClearContents
    oCurRange.Select
    Application.ScreenUpdating = True
End Sub
