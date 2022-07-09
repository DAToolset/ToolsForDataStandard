Option Explicit

Const DICT_ROOT_URL_KO As String = "https://ko.dict.naver.com/"
Const DICT_BASE_URL_KO As String = "https://ko.dict.naver.com/api3/koko/search?query=%s"
Const DICT_ROOT_URL_EN As String = "https://en.dict.naver.com/"
Const DICT_BASE_URL_EN As String = "https://en.dict.naver.com/api3/enko/search?query=%s"

Public Enum DicToSearch
    dtsKorean = 1
    dtsEnglish = 2
    dtsAll = 10
End Enum

Public Type TDicSearchResult
    sWord As String
    sMatchType As String
    sSearchEntry As String
    sMeaning As String
    sLinkURL As String
    sLinkWord As String
    sSynonymList As String
    sAntonymList As String
End Type

Public Function DoDicSearch(aDicToSearch As DicToSearch, aWord As String, _
    bIsMatchTypeExact As Boolean, bIsMatchTypeTermOr As Boolean, bIsMatchTypeAllTerm As Boolean, _
    aMaxResultCount As Long) As TDicSearchResult

    Dim sDicRootURL As String, sBaseURL As String, sURL As String, sURLData As String, sWord As String, oDicSearchResult As TDicSearchResult

    Dim oParsedDic As Dictionary
    Dim oItem As Dictionary, oMeansCollector As Dictionary, oMeans As Dictionary
    Dim oSimWords As Dictionary, oAntWord As Dictionary
    Dim sPOS As String, sMeaning As String, sLinkURL As String, sLinkWord As String
    Dim s유의어 As String, s유의어목록 As String, s반의어 As String, s반의어목록 As String
    Dim sMatchType As String, sSearchEntry As String, sHandleEntry As String

    Select Case aDicToSearch
        Case dtsKorean
            sDicRootURL = DICT_ROOT_URL_KO
            sBaseURL = DICT_BASE_URL_KO
        Case dtsEnglish
            sDicRootURL = DICT_ROOT_URL_EN
            sBaseURL = DICT_BASE_URL_EN
    End Select

    If aWord = "" Then Exit Function
    sWord = URLEncodeUTF8(aWord)
    sURL = Replace(sBaseURL, "%s", sWord)
    sURLData = GetDataFromURL(sURL, "GET", "", "utf-8") 'URL에서 결과 가져오기
    Set oParsedDic = JsonConverter.ParseJson(sURLData) 'JSON결과를 Dictionary로 변환

    Dim lMatchIdx As Long: lMatchIdx = 0
    Dim lResultCount As Long: lResultCount = 0
    For Each oItem In oParsedDic("searchResultMap")("searchResultListMap")("WORD")("items")
        lResultCount = lResultCount + 1
        If (aMaxResultCount <> 0) And (lResultCount > aMaxResultCount) Then Exit For '결과출력 제한개수 초과시 Loop 종료
        s유의어 = "": s반의어 = ""
        lMatchIdx = lMatchIdx + 1
        'If oItem("matchType") <> "exact:entry" Then Exit For

        sHandleEntry = oItem("handleEntry")
        Select Case oItem("matchType")
            Case "exact:entry"
                sLinkWord = sHandleEntry
                sLinkURL = sDicRootURL + oItem("destinationLink")
                If Not bIsMatchTypeExact Then GoTo Continue_InnerFor
            Case "term:or"
                If Not bIsMatchTypeTermOr Then GoTo Continue_InnerFor
            Case "allterm:proximity:1.000000"
                If Not bIsMatchTypeAllTerm Then GoTo Continue_InnerFor
            Case Else
                
        End Select

        sMatchType = sMatchType + IIf(sMatchType = "", "", vbLf) & CStr(lMatchIdx) & ". " & oItem("matchType")
        sSearchEntry = sSearchEntry + IIf(sSearchEntry = "", "", vbLf) & CStr(lMatchIdx) & ". " & sHandleEntry

        For Each oMeansCollector In oItem("meansCollector")
            'Debug.Print "품사: " & oMeansCollector("partOfSpeech")
            sPOS = ""
            If oMeansCollector.Exists("partOfSpeech") Then
                If Not IsNull(oMeansCollector("partOfSpeech")) Then sPOS = oMeansCollector("partOfSpeech")
            End If
            For Each oMeans In oMeansCollector("means")
                'Debug.Print "뜻: " & oMeans("value")
                If oMeans.Exists("value") Then
                    If Not IsNull(oMeans("value")) Then _
                        sMeaning = sMeaning + IIf(sMeaning = "", "", vbLf) & CStr(lMatchIdx) & ". " & IIf(sPOS = "", "", "[" & sPOS & "] ") & RemoveHTML(oMeans("value"))
                End If
            Next oMeans
        Next oMeansCollector
        For Each oSimWords In oItem("similarWordList")
            If oSimWords.Exists("similarWordName") Then _
                s유의어 = s유의어 + IIf(s유의어 = "", "", ", ") & RemoveHTML(oSimWords("similarWordName"))
        Next oSimWords
        If s유의어 <> "" Then _
            s유의어목록 = s유의어목록 & IIf(s유의어목록 = "", "", vbLf) & CStr(lMatchIdx) & ". " & sHandleEntry & ": " & s유의어

        For Each oAntWord In oItem("antonymWordList")
            If oAntWord.Exists("antonymWordName") Then _
                s반의어 = s반의어 + IIf(s반의어 = "", "", ", ") & RemoveHTML(oAntWord("antonymWordName"))
        Next oAntWord
        If s반의어 <> "" Then _
            s반의어목록 = s반의어목록 & IIf(s반의어목록 = "", "", vbLf) & CStr(lMatchIdx) & ". " & sHandleEntry & ": " & s반의어

Continue_InnerFor:
    Next oItem

    If sMeaning = "" Then
        sMeaning = "#NOT FOUND#": sMatchType = sMeaning: sSearchEntry = sMeaning
    End If

    '결과값 반환
    With oDicSearchResult
        .sWord = aWord
        .sMatchType = sMatchType
        .sSearchEntry = sSearchEntry
        .sMeaning = sMeaning
        .sLinkWord = sLinkWord
        .sLinkURL = Replace(sLinkURL, "#", "%23") 'Excel에서 #기호를 내부적으로 #20 - #20 으로 치환하는 것을 방지
        .sSynonymList = s유의어목록
        .sAntonymList = s반의어목록
    End With
    DoDicSearch = oDicSearchResult
End Function

