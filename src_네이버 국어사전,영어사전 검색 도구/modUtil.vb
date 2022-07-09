Option Explicit

Function GetDataFromURL(strURL, strMethod, strPostData, Optional strCharSet = "UTF-8")
  Dim lngTimeout
  Dim strUserAgentString
  Dim intSslErrorIgnoreFlags
  Dim blnEnableRedirects
  Dim blnEnableHttpsToHttpRedirects
  Dim strHostOverride
  Dim strLogin
  Dim strPassword
  Dim strResponseText
  Dim objWinHttp
  lngTimeout = 59000
  strUserAgentString = "http_requester/0.1"
  intSslErrorIgnoreFlags = 13056 ' 13056: ignore all err, 0: accept no err
  blnEnableRedirects = True
  blnEnableHttpsToHttpRedirects = True
  strHostOverride = ""
  strLogin = ""
  strPassword = ""
  Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
  '--------------------------------------------------------------------
  'objWinHttp.SetProxy 2, "xxx.xxx.xxx.xxx:xxxx", "" 'Proxy를 사용하는 환경에서 설정
  '--------------------------------------------------------------------
  objWinHttp.SetTimeouts lngTimeout, lngTimeout, lngTimeout, lngTimeout
  objWinHttp.Open strMethod, strURL
  If strMethod = "POST" Then
    objWinHttp.SetRequestHeader "Content-type", "application/x-www-form-urlencoded; charset=UTF-8"
  Else
    objWinHttp.SetRequestHeader "Content-type", "text/html; charset=euc-kr"
  End If
  If strHostOverride <> "" Then
    objWinHttp.SetRequestHeader "Host", strHostOverride
  End If
  objWinHttp.SetRequestHeader "Referer", "https://ko.dict.naver.com/" '2022-07-08 추가

  objWinHttp.Option(0) = strUserAgentString
  objWinHttp.Option(4) = intSslErrorIgnoreFlags
  objWinHttp.Option(6) = blnEnableRedirects
  objWinHttp.Option(12) = blnEnableHttpsToHttpRedirects
  If (strLogin <> "") And (strPassword <> "") Then
    objWinHttp.SetCredentials strLogin, strPassword, 0
  End If
  On Error Resume Next
  objWinHttp.Send (strPostData)
  objWinHttp.WaitForResponse
  If Err.Number = 0 Then
    If objWinHttp.Status = "200" Then
      'GetDataFromURL = objWinHttp.ResponseText
      GetDataFromURL = BinaryToText(objWinHttp.ResponseBody, strCharSet)
    Else
      GetDataFromURL = "HTTP " & objWinHttp.Status & " " & _
        objWinHttp.StatusText
    End If
  Else
    GetDataFromURL = "Error " & Err.Number & " " & Err.Source & " " & _
      Err.Description
  End If
  On Error GoTo 0
  Set objWinHttp = Nothing
End Function

Function BinaryToText(BinaryData, CharSet)
    Const adTypeText = 2
    Const adTypeBinary = 1
    
    Dim BinaryStream
    Set BinaryStream = CreateObject("ADODB.Stream")
    
    '원본 데이터 타입
    BinaryStream.Type = adTypeBinary
    BinaryStream.Open
    BinaryStream.Write BinaryData
    
    BinaryStream.Position = 0
    BinaryStream.Type = adTypeText
    
    BinaryStream.CharSet = CharSet
    
    BinaryToText = BinaryStream.ReadText
    Set BinaryStream = Nothing
End Function

'Function stripHTML(strHTML) As String
''Strips the HTML tags from strHTML using split and join
'
'  'Ensure that strHTML contains something
'  If Len(strHTML) = 0 Then
'    stripHTML = strHTML
'    Exit Function
'  End If
'
'  Dim arysplit, i, j, strOutput
'
'  arysplit = Split(strHTML, "<")
'
'  'Assuming strHTML is nonempty, we want to start iterating
'  'from the 2nd array postition
'  If Len(arysplit(0)) > 0 Then j = 1 Else j = 0
'
'  'Loop through each instance of the array
'  For i = j To UBound(arysplit)
'     'Do we find a matching > sign?
'     If InStr(arysplit(i), ">") Then
'       'If so, snip out all the text between the start of the string
'       'and the > sign
'       arysplit(i) = Mid(arysplit(i), InStr(arysplit(i), ">") + 1)
'     Else
'       'Ah, the < was was nonmatching
'       arysplit(i) = "<" & arysplit(i)
'     End If
'  Next
'
'  'Rejoin the array into a single string
'  strOutput = Join(arysplit, "")
'
'  'Snip out the first <
'  strOutput = Mid(strOutput, 2 - j)
'
'  'Convert < and > to &lt; and &gt;
'  strOutput = Replace(strOutput, ">", "&gt;")
'  strOutput = Replace(strOutput, "<", "&lt;")
'  strOutput = Replace(strOutput, vbCrLf, " ")
'
'  stripHTML = Trim(strOutput)
'End Function

Function RemoveHTML_old(strText)
    Dim TAGLIST
    TAGLIST = ";!--;!DOCTYPE;A;ACRONYM;ADDRESS;APPLET;AREA;B;BASE;BASEFONT;" & _
              "BGSOUND;BIG;BLOCKQUOTE;BODY;BR;BUTTON;CAPTION;CENTER;CITE;CODE;" & _
              "COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EM;EMBED;FIELDSET;" & _
              "FONT;FORM;FRAME;FRAMESET;HEAD;H1;H2;H3;H4;H5;H6;HR;HTML;I;IFRAME;IMG;" & _
              "INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LI;LINK;LISTING;MAP;MARQUEE;" & _
              "MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OL;OPTION;P;PARAM;PLAINTEXT;" & _
              "PRE;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STRONG;STYLE;SUB;SUP;" & _
              "TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;U;UL;VAR;WBR;XMP;"

    Const BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"
    
    Dim nPos1
    Dim nPos2
    Dim nPos3
    Dim strResult
    Dim strTagName
    Dim bRemove
    Dim bSearchForBlock
    
    nPos1 = InStr(strText, "<")
    Do While nPos1 > 0
        nPos2 = InStr(nPos1 + 1, strText, ">")
        If nPos2 > 0 Then
            strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1)
            strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ")

            nPos3 = InStr(strTagName, " ")
            If nPos3 > 0 Then
                strTagName = Left(strTagName, nPos3 - 1)
            End If
            
            If Left(strTagName, 1) = "/" Then
                strTagName = Mid(strTagName, 2)
                bSearchForBlock = False
            Else
                bSearchForBlock = True
            End If
            
            If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                bRemove = True
                If bSearchForBlock Then
                    If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                        nPos2 = Len(strText)
                        nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare)
                        If nPos3 > 0 Then
                            nPos3 = InStr(nPos3 + 1, strText, ">")
                        End If
                        
                        If nPos3 > 0 Then
                            nPos2 = nPos3
                        End If
                    End If
                End If
            ElseIf strTagName Like "!--*" Then
                bRemove = True
            Else
                bRemove = False
            End If
            
            If bRemove Then
                strResult = strResult & Left(strText, nPos1 - 1)
                strText = Mid(strText, nPos2 + 1)
            Else
                strResult = strResult & Left(strText, nPos1)
                strText = Mid(strText, nPos1 + 1)
            End If
        Else
            strResult = strResult & strText
            strText = ""
        End If
        
        nPos1 = InStr(strText, "<")
    Loop
    strResult = HTMLEntititesDecode(strResult & strText)
    
    RemoveHTML_old = strResult
End Function

'*** 검색결과에서 불필요한 Text 제거 ***

'.+어사전 더보기  -- 정규식 Replace
'0 1 2   -- Replace
'반복속도:   간격없음 --Replace
'간격있음 --Replace
'간격길게 --Replace
'[\r\n?|\n]{2,}  -- 정규식 Replace, 2줄 이상 행분리를 한행으로
'(.+어사전)  --> __________\r\n$1
Public Function TidyText(aText As String) As String
    Dim sResult As String
    sResult = aText
'    sResult = RegExReplace(sResult, ".+어사전 더보기", "")
'    sResult = Replace(sResult, "0 1 2", "")
'    sResult = Replace(sResult, "반복속도 :  간격없음", "")
'    sResult = Replace(sResult, "간격있음", "")
'    sResult = Replace(sResult, "간격길게", "")
'    sResult = RegExReplace(sResult, "[\r\n?|\n]{2,}", vbLf)
'    sResult = RegExReplace(sResult, "(^.{0,15}어사전)", "__________" + vbLf + "$1")
'    sResult = RegExReplace(sResult, "(인도네시아사전)", "__________" + vbLf + "$1")
'
'    sResult = Replace(sResult, "발음재생", "")
'    sResult = Replace(sResult, "단어장 저장", "")
'    sResult = Replace(sResult, "오픈사전", "")
'    sResult = Replace(sResult, "도움말", "")
'    sResult = Replace(sResult, "단어 더보기", "")
'    sResult = Replace(sResult, "단어/숙어 더보기", "")

    sResult = RegExReplace(sResult, "<sup>[^>]+>", "") '<sup>...</sup> 삭제
    TidyText = sResult
End Function

Function RemoveHTML(sString As String) As String
    On Error GoTo Error_Handler
    Dim oRegEx As Object, sString2 As String

    sString2 = RegExReplace(sString, "<sup>[^>]+>", "") '<sup>...</sup> 삭제
    Set oRegEx = CreateObject("vbscript.regexp")
 
    With oRegEx
        'Patterns see: http://regexlib.com/Search.aspx?k=html%20tags
        .Pattern = "<[^>]+>"    'basic html pattern
        '.Pattern = "<!*[^<>]*>"    'html tags and comments
        .Global = True
        .IgnoreCase = True
        '.MultiLine = True
        .MultiLine = False
    End With

    RemoveHTML = HTMLEntititesDecode(oRegEx.Replace(sString2, ""))
    'RemoveHTML = TidyText(RemoveHTML)

Error_Handler_Exit:
    On Error Resume Next
    Set oRegEx = Nothing
    Exit Function

Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: RemoveHTML" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Function RegExReplace(sString As String, sPattern As String, sReplaceStrring As String) As String
    On Error GoTo Error_Handler
    Dim oRegEx          As Object
 
    Set oRegEx = CreateObject("vbscript.regexp")
 
    With oRegEx
        .Pattern = sPattern
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
    End With

    RegExReplace = oRegEx.Replace(sString, sReplaceStrring)

Error_Handler_Exit:
    On Error Resume Next
    Set oRegEx = Nothing
    Exit Function

Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: RegExReplace" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Public Function URLDecodeUTF8(ByVal pURL)
    Dim i, s1, s2, s3, u1, u2, Result
    pURL = Replace(pURL, "+", " ")

    For i = 1 To Len(pURL)
        If Mid(pURL, i, 1) = "%" Then
            s1 = CLng("&H" & Mid(pURL, i + 1, 2))
    
            '2바이트일 경우
            If ((s1 And &HC0) = &HC0) And ((s1 And &HE0) <> &HE0) Then
               s2 = CLng("&H" & Mid(pURL, i + 4, 2))
            
               u1 = (s1 And &H1C) / &H4
               u2 = ((s1 And &H3) * &H4 + ((s2 And &H30) / &H10)) * &H10
               u2 = u2 + (s2 And &HF)
               Result = Result & ChrW((u1 * &H100) + u2)
               i = i + 5
            '3바이트일 경우
            ElseIf (s1 And &HE0 = &HE0) Then
              s2 = CLng("&H" & Mid(pURL, i + 4, 2))
              s3 = CLng("&H" & Mid(pURL, i + 7, 2))
              u1 = ((s1 And &HF) * &H10)
              u1 = u1 + ((s2 And &H3C) / &H4)
              u2 = ((s2 And &H3) * &H4 + (s3 And &H30) / &H10) * &H10
              u2 = u2 + (s3 And &HF)
              Result = Result & ChrW((u1 * &H100) + u2)
              i = i + 8
            End If
        Else
            Result = Result & Mid(pURL, i, 1)
        End If
    Next
    URLDecodeUTF8 = Result
End Function


Public Function URLEncodeUTF8_old(ByVal szSource)

Dim szChar, WideChar, nLength, i, Result
nLength = Len(szSource)

'szSource = Replace(szSource," ","+")

For i = 1 To nLength
 szChar = Mid(szSource, i, 1)

 If Asc(szChar) < 0 Then
  WideChar = CLng(AscB(MidB(szChar, 2, 1))) * 256 + AscB(MidB(szChar, 1, 1))

  If (WideChar And &HFF80) = 0 Then
   Result = Result & "%" & Hex(WideChar)
  ElseIf (WideChar And &HF000) = 0 Then
   Result = Result & _
   "%" & Hex(CInt((WideChar And &HFFC0) / 64) Or &HC0) & _
   "%" & Hex(WideChar And &H3F Or &H80)
  Else
   Result = Result & _
   "%" & Hex(CInt((WideChar And &HF000) / 4096) Or &HE0) & _
   "%" & Hex(CInt((WideChar And &HFFC0) / 64) And &H3F Or &H80) & _
   "%" & Hex(WideChar And &H3F Or &H80)
  End If
 Else
  Result = Result + szChar
 End If
Next
URLEncodeUTF8_old = Result
End Function

Public Function URLEncodeUTF8( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String
  Dim bytes() As Byte, b As Byte, i As Integer, space As String

  If SpaceAsPlus Then space = "+" Else space = "%20"

  If Len(StringVal) > 0 Then
    With New ADODB.Stream
      .Mode = adModeReadWrite
      .Type = adTypeText
      .CharSet = "UTF-8"
      .Open
      .WriteText StringVal
      .Position = 0
      .Type = adTypeBinary
      .Position = 3 ' skip BOM
      bytes = .Read
    End With

    ReDim Result(UBound(bytes)) As String

    For i = UBound(bytes) To 0 Step -1
      b = bytes(i)
      Select Case b
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          Result(i) = Chr(b)
        Case 32
          Result(i) = space
        Case 0 To 15
          Result(i) = "%0" & Hex(b)
        Case Else
          Result(i) = "%" & Hex(b)
      End Select
    Next i

    URLEncodeUTF8 = Join(Result, "")
  End If
End Function

Public Function 한글제거(strData As String) As String
    Dim i As Integer
    Dim intLen As Integer
    Dim strX As String
    Dim strTemp As String
    
    intLen = Len(strData)
    
    For i = 1 To intLen
        strX = Mid$(strData, i, 1)
        Select Case strX
            Case "ㄱ" To "ㅎ", "가" To "힣"
            Case Else
                strTemp = strTemp & strX
        End Select
    Next i
    
    한글제거 = Trim(strTemp)

End Function

Public Function IsAlNum(strData As String) As String
    Dim i As Integer
    Dim intLen As Integer
    Dim strX As String
    Dim strTemp As String

    intLen = Len(strData)

    For i = 1 To intLen
        strX = Mid$(strData, i, 1)
        Select Case strX
            Case "a" To "z", "A" To "Z"
            Case INList(strX, "`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "-", "=", "[", "]", "\", "{", "}", "|", ";", "'", "'", ":", """, """, ",", ".", "/", "<", ">", "?", ",", " ")
            Case Else
                strTemp = strTemp & strX
        End Select
    Next i

    IsAlNum = (strTemp = "")

End Function


Public Function TrimLeadingSpace(sData As String) As String
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String

    'strPattern = "^\s+"
    strPattern = "^[\s\xA0]+|[\s\xA0]+$"

    If strPattern <> "" Then
        strInput = sData
        strReplace = vbLf
        'strReplace = ""

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
            TrimLeadingSpace = Replace(regEx.Replace(strInput, strReplace), vbLf + vbLf, vbLf)
            TrimLeadingSpace = Replace(TrimLeadingSpace, vbLf + vbLf, vbLf)
            TrimLeadingSpace = Mid(TrimLeadingSpace, 2, Len(TrimLeadingSpace) - 2)
        Else
            TrimLeadingSpace = sData
        End If
    End If
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'List 값중 하나라도 포함되어 있는 경우 true를 retrun함
'-- aValue: 비교할 값
'-- aList: 비교대상 List
'-- 예1: INList("X", "X", "Y", "Z") : true ==> X는 X, Y, Z 에 포함되어 있음.
'-- 예2: INList("A", "X", "Y", "Z") : false ==> A는 X, Y, Z 에 포함되어 있지 않음.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function INList(aValue As String, ParamArray aList() As Variant) As Boolean
    Dim bOut As Boolean, i As Integer

    For i = 0 To UBound(aList)
        bOut = IIf(i = 0, aValue = aList(i), bOut Or aValue = aList(i))
        If bOut = True Then Exit For '일치하는 값을 찾으면 이후는 더 찾을 필요 없음.
    Next i

    INList = bOut
End Function


Public Function HTMLEntititesDecode(p_strText As String) As String
    Dim strTemp As String
    strTemp = p_strText
    strTemp = Replace(strTemp, "&quot;", """")
    strTemp = Replace(strTemp, "&amp;", "&")
    strTemp = Replace(strTemp, "&apos;", "'")
    strTemp = Replace(strTemp, "&lt;", "<")
    strTemp = Replace(strTemp, "&gt;", ">")
    strTemp = Replace(strTemp, "&nbsp;", "")
    strTemp = Replace(strTemp, "&iexcl;", "¡")
    strTemp = Replace(strTemp, "&cent;", "￠")
    strTemp = Replace(strTemp, "&pound;", "￡")
    strTemp = Replace(strTemp, "&curren;", "¤")
    strTemp = Replace(strTemp, "&yen;", "￥")
    strTemp = Replace(strTemp, "&brvbar;", "|")
    strTemp = Replace(strTemp, "&sect;", "§")
    strTemp = Replace(strTemp, "&uml;", "¨")
    strTemp = Replace(strTemp, "&copy;", "ⓒ")
    strTemp = Replace(strTemp, "&ordf;", "ª")
    strTemp = Replace(strTemp, "&laquo;", "≪")
    strTemp = Replace(strTemp, "&not;", "￢")
    strTemp = Replace(strTemp, "*", "")
    strTemp = Replace(strTemp, "&reg;", "®")
    strTemp = Replace(strTemp, "&macr;", "?")
    strTemp = Replace(strTemp, "&deg;", "°")
    strTemp = Replace(strTemp, "&plusmn;", "±")
    strTemp = Replace(strTemp, "&sup2;", "²")
    strTemp = Replace(strTemp, "&sup3;", "³")
    strTemp = Replace(strTemp, "&acute;", "´")
    strTemp = Replace(strTemp, "&micro;", "μ")
    strTemp = Replace(strTemp, "&para;", "¶")
    strTemp = Replace(strTemp, "&middot;", "·")
    strTemp = Replace(strTemp, "&cedil;", "¸")
    strTemp = Replace(strTemp, "&sup1;", "¹")
    strTemp = Replace(strTemp, "&ordm;", "º")
    strTemp = Replace(strTemp, "&raquo;", "≫")
    strTemp = Replace(strTemp, "&frac14;", "¼")
    strTemp = Replace(strTemp, "&frac12;", "½")
    strTemp = Replace(strTemp, "&frac34;", "¾")
    strTemp = Replace(strTemp, "&iquest;", "¿")
    strTemp = Replace(strTemp, "&Agrave;", "A")
    strTemp = Replace(strTemp, "&Aacute;", "A")
    strTemp = Replace(strTemp, "&Acirc;", "A")
    strTemp = Replace(strTemp, "&Atilde;", "A")
    strTemp = Replace(strTemp, "&Auml;", "A")
    strTemp = Replace(strTemp, "&Aring;", "A")
    strTemp = Replace(strTemp, "&AElig;", "Æ")
    strTemp = Replace(strTemp, "&Ccedil;", "C")
    strTemp = Replace(strTemp, "&Egrave;", "E")
    strTemp = Replace(strTemp, "&Eacute;", "E")
    strTemp = Replace(strTemp, "&Ecirc;", "E")
    strTemp = Replace(strTemp, "&Euml;", "E")
    strTemp = Replace(strTemp, "&Igrave;", "I")
    strTemp = Replace(strTemp, "&Iacute;", "I")
    strTemp = Replace(strTemp, "&Icirc;", "I")
    strTemp = Replace(strTemp, "&Iuml;", "I")
    strTemp = Replace(strTemp, "&ETH;", "Ð")
    strTemp = Replace(strTemp, "&Ntilde;", "N")
    strTemp = Replace(strTemp, "&Ograve;", "O")
    strTemp = Replace(strTemp, "&Oacute;", "O")
    strTemp = Replace(strTemp, "&Ocirc;", "O")
    strTemp = Replace(strTemp, "&Otilde;", "O")
    strTemp = Replace(strTemp, "&Ouml;", "O")
    strTemp = Replace(strTemp, "&times;", "×")
    strTemp = Replace(strTemp, "&Oslash;", "Ø")
    strTemp = Replace(strTemp, "&Ugrave;", "U")
    strTemp = Replace(strTemp, "&Uacute;", "U")
    strTemp = Replace(strTemp, "&Ucirc;", "U")
    strTemp = Replace(strTemp, "&Uuml;", "U")
    strTemp = Replace(strTemp, "&Yacute;", "Y")
    strTemp = Replace(strTemp, "&THORN;", "Þ")
    strTemp = Replace(strTemp, "&szlig;", "ß")
    strTemp = Replace(strTemp, "&agrave;", "a")
    strTemp = Replace(strTemp, "&aacute;", "a")
    strTemp = Replace(strTemp, "&acirc;", "a")
    strTemp = Replace(strTemp, "&atilde;", "a")
    strTemp = Replace(strTemp, "&auml;", "a")
    strTemp = Replace(strTemp, "&aring;", "a")
    strTemp = Replace(strTemp, "&aelig;", "æ")
    strTemp = Replace(strTemp, "&ccedil;", "c")
    strTemp = Replace(strTemp, "&egrave;", "e")
    strTemp = Replace(strTemp, "&eacute;", "e")
    strTemp = Replace(strTemp, "&ecirc;", "e")
    strTemp = Replace(strTemp, "&euml;", "e")
    strTemp = Replace(strTemp, "&igrave;", "i")
    strTemp = Replace(strTemp, "&iacute;", "i")
    strTemp = Replace(strTemp, "&icirc;", "i")
    strTemp = Replace(strTemp, "&iuml;", "i")
    strTemp = Replace(strTemp, "&eth;", "ð")
    strTemp = Replace(strTemp, "&ntilde;", "n")
    strTemp = Replace(strTemp, "&ograve;", "o")
    strTemp = Replace(strTemp, "&oacute;", "o")
    strTemp = Replace(strTemp, "&ocirc;", "o")
    strTemp = Replace(strTemp, "&otilde;", "o")
    strTemp = Replace(strTemp, "&ouml;", "o")
    strTemp = Replace(strTemp, "&divide;", "÷")
    strTemp = Replace(strTemp, "&oslash;", "ø")
    strTemp = Replace(strTemp, "&ugrave;", "u")
    strTemp = Replace(strTemp, "&uacute;", "u")
    strTemp = Replace(strTemp, "&ucirc;", "u")
    strTemp = Replace(strTemp, "&uuml;", "u")
    strTemp = Replace(strTemp, "&yacute;", "y")
    strTemp = Replace(strTemp, "&thorn;", "þ")
    strTemp = Replace(strTemp, "&yuml;", "y")
    strTemp = Replace(strTemp, "&OElig;", "Œ")
    strTemp = Replace(strTemp, "&oelig;", "œ")
    strTemp = Replace(strTemp, "&Scaron;", "?")
    strTemp = Replace(strTemp, "&scaron;", "?")
    strTemp = Replace(strTemp, "&Yuml;", "?")
    strTemp = Replace(strTemp, "&fnof;", "?")
    strTemp = Replace(strTemp, "&circ;", "?")
    strTemp = Replace(strTemp, "&tilde;", "?")
    strTemp = Replace(strTemp, "&thinsp;", "")
    strTemp = Replace(strTemp, "&zwnj;", "")
    strTemp = Replace(strTemp, "&zwj;", "")
    strTemp = Replace(strTemp, "&lrm;", "")
    strTemp = Replace(strTemp, "&rlm;", "")
    strTemp = Replace(strTemp, "&ndash;", "?")
    strTemp = Replace(strTemp, "&mdash;", "?")
    strTemp = Replace(strTemp, "&lsquo;", "‘")
    strTemp = Replace(strTemp, "&rsquo;", "’")
    strTemp = Replace(strTemp, "&sbquo;", "?")
    strTemp = Replace(strTemp, "&ldquo;", "“")
    strTemp = Replace(strTemp, "&rdquo;", "”")
    strTemp = Replace(strTemp, "&bdquo;", "?")
    strTemp = Replace(strTemp, "&dagger;", "†")
    strTemp = Replace(strTemp, "&Dagger;", "‡")
    strTemp = Replace(strTemp, "&bull;", "?")
    strTemp = Replace(strTemp, "&hellip;", "…")
    strTemp = Replace(strTemp, "&permil;", "‰")
    strTemp = Replace(strTemp, "&lsaquo;", "?")
    strTemp = Replace(strTemp, "&rsaquo;", "?")
    strTemp = Replace(strTemp, "&euro;", "€")
    strTemp = Replace(strTemp, "&trade;", "™")
    strTemp = Replace(strTemp, "&frasl;", "/")
    strTemp = Replace(strTemp, "&#124;", Chr(124))
    HTMLEntititesDecode = strTemp
End Function


Public Function IsAlphaBetical(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String

    'returns true if all characters in a string are alphabetical
    'returns false otherwise or for empty string

    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[A-Za-z]" Then Exit Function
        Next
    
    IsAlphaBetical = True
    End If

End Function

Public Function IsNumericOnly(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String

    'returns true if all characters in string are numeric
    'returns false otherwise or for empty string

    'this is different than VB's isNumeric
    'isNumeric returns true for something like 90.09
    'This function will return false

    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9]" Then Exit Function
        Next

    IsNumericOnly = True
    End If

End Function

Public Function IsAlphaNumeric(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String

    'returns true if all characters in a string are alphabetical
    '   or numeric
    'returns false otherwise or for empty string

    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9A-Za-z]" Then Exit Function
        Next

    IsAlphaNumeric = True
    End If

End Function

Public Function IsAlphaNumericSpecial(TestString As String) As Boolean
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String

    'returns true if all characters in a string are alphabetical
    '   or numeric
    'returns false otherwise or for empty string

    IsAlphaNumericSpecial = False
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            'If Not sChar Like "[!-~]" Then Exit Function
            If Not sChar Like "[0-9A-Za-z]" And Not INList(sChar, "`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "-", "=", "[", "]", "\", "{", "}", "|", ";", "'", "'", ":", """, """, ",", ".", "/", "<", ">", "?", ",", " ") Then Exit Function
        Next

    IsAlphaNumericSpecial = True
    End If

End Function

