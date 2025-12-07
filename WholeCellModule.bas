Attribute VB_Name = "WholeCellModule"
'===== Module: WholeCellModule ============================================

Option Explicit

' ========================= WholeCell API (Mac + Windows) =========================
' Public functions:
'   WholeCellFetch(imeiOrEsn As String) As Object
'   WholeCellCost(imeiOrEsn As String) As Variant
' ================================================================================

Private Const DEBUG_MODE As Boolean = False

Private Const API_BASE_URL As String = "https://api.wholecell.io/api/v1/inventories"
Private Const API_APP_ID As String = "yBsKJYb5WvL_V-2LCNuObgNR4t8iBxZ36S34vDEFsAVA"
Private Const API_APP_SECRET As String = "rjXwBWktyT26QNbfOCIRsgn8E2QCgIioy-l00R7f7oWA"

' --------- Public API ---------

Public Function WholeCellCost(imeiOrEsn As String) As Variant
    Dim r As Object: Set r = WholeCellFetch(imeiOrEsn)
    If Not r Is Nothing And r.Exists("success") And r("success") Then
        WholeCellCost = r("universalUsd")
    Else
        WholeCellCost = CVErr(xlErrNA)
    End If
End Function

Public Function WholeCellFetch(imeiOrEsn As String) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    result("success") = False
    result("universalUsd") = 0#
    result("model") = ""
    result("grade") = ""
    result("esn") = ""
    result("error") = ""

    imeiOrEsn = Trim$(imeiOrEsn)
    If Len(imeiOrEsn) = 0 Then
        result("error") = "No IMEI/ESN provided"
        Set WholeCellFetch = result
        Exit Function
    End If

    ' Keep Excel responsive
    DoEvents

    Dim authB64 As String
    authB64 = Base64Encode(API_APP_ID & ":" & API_APP_SECRET)

    Dim json As String
    Dim apiUrl As String
    apiUrl = API_BASE_URL & "?esn=" & imeiOrEsn
    json = HttpGet(apiUrl, authB64, API_APP_ID)

    If Len(json) = 0 Then
        result("error") = "Failed to connect to WholeCell API"
        Set WholeCellFetch = result
        Exit Function
    End If
   
    If DEBUG_MODE Then
        Debug.Print "API URL: " & apiUrl
        Debug.Print "API Response Length: " & Len(json)
        Debug.Print "First 500 chars: " & Left(json, 500)
        Debug.Print "Searching for IMEI/ESN: " & imeiOrEsn
    End If

    If ParseOneInventoryItem(json, imeiOrEsn, result) Then
        result("success") = True
        result("error") = ""
    Else
        result("error") = "Device not found"
    End If

    ' Keep Excel responsive after processing
    DoEvents
    
    Set WholeCellFetch = result
End Function

Public Sub TestWholeCellConnection()
    Dim testIMEI As String
    testIMEI = InputBox("Enter a test IMEI/ESN:", "Test WholeCell Connection")
    If Len(testIMEI) = 0 Then Exit Sub

    Dim r As Object: Set r = WholeCellFetch(testIMEI)
    If r("success") Then
        MsgBox "Model: " & r("model") & vbCrLf & _
               "Grade: " & r("grade") & vbCrLf & _
               "Universal Cost: $" & Format(r("universalUsd"), "#,##0.00") & vbCrLf & _
               "ESN: " & r("esn"), vbInformation, "WholeCell Connection Test"
    Else
        MsgBox "Failed to retrieve device: " & r("error"), vbCritical, "WholeCell Connection Test"
    End If
End Sub

Public Sub TestKnownIMEIs()
    ' Test with known IMEIs from the test file
    Dim testIMEIs As Variant
    testIMEIs = Array("H95DHMF9Q1GC", "F9FG5XAJQ1GC", "F9GG5BXXQ1GC")
   
    Dim i As Long, r As Object, results As String
    results = "Test Results for Known IMEIs:" & vbCrLf & vbCrLf
   
    For i = 0 To UBound(testIMEIs)
        Set r = WholeCellFetch(CStr(testIMEIs(i)))
        results = results & "IMEI: " & testIMEIs(i) & vbCrLf
       
        If r("success") Then
            results = results & "  ? Found: " & r("model") & " (Grade " & r("grade") & ") - $" & Format(r("universalUsd"), "#,##0.00") & vbCrLf
        Else
            results = results & "  ? Error: " & r("error") & vbCrLf
        End If
        results = results & vbCrLf
    Next i
   
    MsgBox results, vbInformation, "Known IMEI Test Results"
End Sub

' --------- HTTP (cross-platform) ---------

Private Function HttpGet(url As String, authB64 As String, appId As String) As String
    If IsMac() Then
        HttpGet = MacCurlGet(url, authB64, appId)
    Else
        HttpGet = WinHttpGet(url, authB64, appId)
    End If
End Function

Private Function WinHttpGet(url As String, authB64 As String, appId As String) As String
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    If http Is Nothing Then Exit Function

    ' Keep Excel responsive before network call
    DoEvents
    
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Basic " & authB64
    http.setRequestHeader "X-App-Id", appId
    http.setRequestHeader "Accept", "application/json"
    http.send
    
    ' Keep Excel responsive after network call
    DoEvents
    
    If http.Status = 200 Then WinHttpGet = http.responseText
End Function

Private Function MacCurlGet(url As String, authB64 As String, appId As String) As String
    On Error GoTo FailFast
    Dim cmd As String
    cmd = "/usr/bin/curl -sS -L " & _
          "-H " & AppleScriptQuote("Authorization: Basic " & authB64) & " " & _
          "-H " & AppleScriptQuote("X-App-Id: " & appId) & " " & _
          "-H " & AppleScriptQuote("Accept: application/json") & " " & _
          AppleScriptQuote(url)

    ' Keep Excel responsive before network call
    DoEvents
    
    MacCurlGet = MacScript("do shell script " & AppleScriptQuote(cmd))
    
    ' Keep Excel responsive after network call
    DoEvents
    Exit Function
FailFast:
    MacCurlGet = ""
End Function

' --------- Parsing helpers (lightweight text scan) ---------

Private Function ParseOneInventoryItem(json As String, searchEsn As String, ByRef result As Object) As Boolean
    Dim pos As Long, esnVal As String, itemStart As Long, itemEnd As Long, item As String
    searchEsn = UCase$(Trim$(searchEsn))
   
    If DEBUG_MODE Then
        Debug.Print "ParseOneInventoryItem: Looking for " & searchEsn
        Dim esnCount As Long: esnCount = 0
    End If

    pos = 1
    Do
        ' Keep Excel responsive during JSON parsing
        DoEvents
        
        pos = InStr(pos, json, """esn""")
        If pos = 0 Then Exit Do

        esnVal = ExtractStringValue(json, pos)
        If DEBUG_MODE Then
            esnCount = esnCount + 1
            Debug.Print "Found ESN #" & esnCount & ": " & esnVal
        End If
       
        If UCase$(Trim$(esnVal)) = searchEsn Then
            itemStart = FindObjectStart(json, pos)
            itemEnd = FindObjectEnd(json, pos)
            If itemStart > 0 And itemEnd > itemStart Then
                item = Mid$(json, itemStart, itemEnd - itemStart + 1)

                result("esn") = esnVal
                result("universalUsd") = ExtractNumberValue(item, "total_price_paid") / 100#
                result("grade") = ExtractNestedString(item, "product_variation", "grade")

                Dim prod As String
                prod = ExtractSection(item, """product""")
                result("model") = BuildModelString( _
                                    ExtractStringAfterKey(prod, "manufacturer"), _
                                    ExtractStringAfterKey(prod, "model"), _
                                    ExtractStringAfterKey(prod, "capacity"), _
                                    ExtractStringAfterKey(prod, "color") _
                                  )
                ParseOneInventoryItem = True
                Exit Function
            End If
        End If
        pos = pos + 5
    Loop
   
    If DEBUG_MODE Then
        Debug.Print "ParseOneInventoryItem: Search completed. Total ESNs found: " & esnCount
        Debug.Print "ParseOneInventoryItem: No match found for " & searchEsn
    End If
End Function

Private Function ExtractStringValue(json As String, afterPos As Long) As String
    Dim c As Long, q1 As Long, q2 As Long
    c = InStr(afterPos, json, ":"): If c = 0 Then Exit Function
    q1 = InStr(c, json, """"):     If q1 = 0 Then Exit Function
    q2 = InStr(q1 + 1, json, """"): If q2 = 0 Then Exit Function
    ExtractStringValue = Mid$(json, q1 + 1, q2 - q1 - 1)
End Function

Private Function ExtractStringAfterKey(js As String, key As String) As String
    ExtractStringAfterKey = ExtractStringValue(js, InStr(1, js, """" & key & """"))
End Function

Private Function ExtractNestedString(js As String, parentKey As String, childKey As String) As String
    Dim sec As String
    sec = ExtractSection(js, """" & parentKey & """")
    If Len(sec) > 0 Then
        ExtractNestedString = ExtractStringAfterKey(sec, childKey)
    Else
        ExtractNestedString = ""
    End If
End Function

Private Function ExtractNumberValue(js As String, key As String) As Double
    Dim k As Long, c As Long, s As Long, e As Long, ch As String
    k = InStr(1, js, """" & key & """"): If k = 0 Then Exit Function
    c = InStr(k, js, ":"):               If c = 0 Then Exit Function

    For s = c + 1 To Len(js)
        ch = Mid$(js, s, 1)
        If ch Like "[0-9.-]" Then Exit For
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Function
    Next s

    For e = s To Len(js)
        ch = Mid$(js, e, 1)
        If Not (ch Like "[0-9.-]") Then Exit For
    Next e

    If e > s Then ExtractNumberValue = Val(Mid$(js, s, e - s))
End Function

Private Function ExtractSection(js As String, startMarker As String) As String
    Dim p As Long, b As Long, i As Long, ch As String, depth As Long
    p = InStr(1, js, startMarker): If p = 0 Then Exit Function
    b = InStr(p, js, "{"):         If b = 0 Then Exit Function

    depth = 0
    For i = b To Len(js)
        ch = Mid$(js, i, 1)
        If ch = "{" Then
            depth = depth + 1
        ElseIf ch = "}" Then
            depth = depth - 1
            If depth = 0 Then
                ExtractSection = Mid$(js, b, i - b + 1)
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FindObjectStart(js As String, fromPos As Long) As Long
    Dim i As Long
    For i = fromPos To 1 Step -1
        If Mid$(js, i, 1) = "{" Then FindObjectStart = i: Exit Function
    Next i
End Function

Private Function FindObjectEnd(js As String, fromPos As Long) As Long
    Dim startPos As Long, i As Long, depth As Long, ch As String, inString As Boolean

    startPos = FindObjectStart(js, fromPos): If startPos = 0 Then Exit Function

    For i = startPos To Len(js)
        ch = Mid$(js, i, 1)

        If ch = """" Then
            ' toggle only if the quote is not escaped by backslash
            If i = 1 Or Mid$(js, i - 1, 1) <> Chr$(92) Then
                inString = Not inString
            End If
        End If

        If Not inString Then
            If ch = "{" Then depth = depth + 1
            If ch = "}" Then
                depth = depth - 1
                If depth = 0 Then FindObjectEnd = i: Exit Function
            End If
        End If
    Next i
End Function

Private Function BuildModelString(manu As String, mdl As String, cap As String, col As String) As String
    Dim parts As Collection: Set parts = New Collection
    On Error Resume Next
    ' Removed manufacturer from model string - only include model, capacity, and color
    If Len(Trim$(mdl)) > 0 Then parts.Add Trim$(mdl)
    If Len(Trim$(cap)) > 0 Then parts.Add Trim$(cap)
    If Len(Trim$(col)) > 0 Then parts.Add Trim$(col)
    On Error GoTo 0

    Dim i As Long, s As String
    For i = 1 To parts.count
        If i > 1 Then s = s & " "
        s = s & parts(i)
    Next i
    BuildModelString = s
End Function

' --------- Utilities ---------

Private Function IsMac() As Boolean
    IsMac = (InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0)
End Function

Private Function AppleScriptQuote(ByVal s As String) As String
    AppleScriptQuote = """" & Replace(s, """", "\""") & """"
End Function

' SAFE Base64 (no IIf branch-eval problems)
Private Function Base64Encode(text As String) As String
    Const B64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim i As Long, a As Long, b As Long, c As Long
    Dim e1 As Long, e2 As Long, e3 As Long, e4 As Long
    Dim out As String, n As Long

    n = Len(text)
    For i = 1 To n Step 3
        a = Asc(Mid$(text, i, 1))
        If i + 1 <= n Then b = Asc(Mid$(text, i + 1, 1)) Else b = -1
        If i + 2 <= n Then c = Asc(Mid$(text, i + 2, 1)) Else c = -1

        e1 = a \ 4

        If b >= 0 Then
            e2 = ((a And 3) * 16) + (b \ 16)
        Else
            e2 = ((a And 3) * 16)
        End If

        If b >= 0 Then
            If c >= 0 Then
                e3 = ((b And 15) * 4) + (c \ 64)
            Else
                e3 = ((b And 15) * 4)
            End If
        Else
            e3 = 64 ' '='
        End If

        If c >= 0 Then
            e4 = c And 63
        Else
            e4 = 64 ' '='
        End If

        out = out & Mid$(B64, e1 + 1, 1) & _
                    Mid$(B64, e2 + 1, 1) & _
                    IIf(e3 = 64, "=", Mid$(B64, e3 + 1, 1)) & _
                    IIf(e4 = 64, "=", Mid$(B64, e4 + 1, 1))
    Next i

    Base64Encode = out
End Function

'=========================================================================




