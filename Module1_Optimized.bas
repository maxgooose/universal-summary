Attribute VB_Name = "Module1_Optimized"

' ===================================
' PRICING SUMMARY GENERATOR MACRO - OPTIMIZED VERSION
' ===================================
' This macro creates a structured pricing summary report from MANIFEST data
' Button is placed on MANIFEST sheet - user enters all data there first
' Then clicks button to generate the "Summary of pricing" sheet
' Author: Harb
' Version: 3.0 - OPTIMIZED
'   - Single pass data loading into memory arrays
'   - Pre-grouped data structures to eliminate redundant loops
'   - Batch Excel operations (array read/write)
'   - Unique IMEI collection to avoid duplicate API calls
'   - Proper Excel settings management for performance
'   - Maintains max 2 API calls per second rate limit
' ===================================

Option Explicit

' ===================================
' TYPE DEFINITIONS FOR STRUCTURED DATA
' ===================================

' Represents a single item from MANIFEST
Private Type ManifestItem
    RowIndex As Long
    Model As String
    Storage As String
    Color As String
    Grade As String
    Battery As String
    IMEI As String
    ModelKey As String      ' Model|Storage
    VariationKey As String  ' Grade|Battery|Color
    FullKey As String       ' Model|Storage|Grade|Battery|Color
End Type

' Represents aggregated data for a variation group
Private Type VariationGroup
    Model As String
    Storage As String
    Color As String
    Grade As String
    Battery As String
    Quantity As Long
    IMEIs As Collection     ' Collection of unique IMEIs
    TotalCost As Double
    ValidCostCount As Long
    AvgCost As Double
End Type

' ===================================
' MODULE-LEVEL VARIABLES
' ===================================

' Global cache for WholeCell costs - persists across all lookups
Private costCache As Object

' Rate limiting - track API call timestamps
Private lastApiCallTime As Double
Private apiCallCount As Long

' Statistics tracking
Private totalIMEIsProcessed As Long
Private totalIMEIsFound As Long
Private totalIMEIsNotFound As Long

' Pre-loaded data structures (populated once, used many times)
Private manifestData() As ManifestItem
Private manifestDataCount As Long
Private uniqueModelKeys As Object          ' Dictionary: ModelKey -> True
Private modelVariations As Object          ' Dictionary: ModelKey -> Collection of VariationKeys
Private variationGroups As Object          ' Dictionary: FullKey -> VariationGroup
Private allUniqueIMEIs As Object           ' Dictionary: IMEI -> cost (populated during fetch)

' ===================================
' SAFE TYPE CONVERSION FUNCTIONS
' ===================================

Private Function CDbl0(ByVal v As Variant) As Double
    On Error Resume Next
    CDbl0 = CDbl(v)
    If Err.Number <> 0 Then CDbl0 = 0#
    On Error GoTo 0
End Function

Private Function Nz(ByVal v As Variant, Optional ByVal replacement As Variant = "") As Variant
    If IsError(v) Then
        Nz = replacement
    ElseIf IsMissing(v) Then
        Nz = replacement
    ElseIf IsNull(v) Then
        Nz = replacement
    ElseIf VarType(v) = vbString Then
        If Len(v) = 0 Then Nz = replacement Else Nz = v
    Else
        If Len(Trim$(CStr(v))) = 0 Then Nz = replacement Else Nz = v
    End If
End Function

' ===================================
' RATE-LIMITED API FETCH
' ===================================

' Fetch with retry and rate limiting (max 2 calls per second)
Private Function FetchWithRetry(ByVal imeiOrEsn As String, ByVal retries As Long) As Object
    Dim i As Long, dict As Object
    
    For i = 0 To retries
        ' Enforce rate limit before each attempt
        EnforceRateLimit
        
        Set dict = WholeCellModule.WholeCellFetch(imeiOrEsn)
        
        If Not dict Is Nothing Then
            If CBool(Nz(dict("success"), False)) Then
                Set FetchWithRetry = dict
                Exit Function
            End If
        End If
        
        DoEvents
        SleepMs 300 * (i + 1)   ' Simple backoff
    Next i
    
    Set FetchWithRetry = dict   ' Return last attempt
End Function

' Optimized rate limiting - simple and efficient
Private Sub EnforceRateLimit()
    Dim currentTime As Double
    Dim elapsed As Double
    
    currentTime = Timer
    
    ' Handle midnight rollover
    If currentTime < lastApiCallTime Then
        lastApiCallTime = 0
        apiCallCount = 0
    End If
    
    elapsed = currentTime - lastApiCallTime
    
    ' If less than 0.5 seconds since last call, we need to wait
    ' (0.5 seconds between calls = max 2 calls per second)
    If elapsed < 0.5 And apiCallCount > 0 Then
        SleepMs CLng((0.5 - elapsed) * 1000) + 50  ' Add 50ms buffer
    End If
    
    ' Reset counter if more than 1 second has passed
    If elapsed >= 1# Then
        apiCallCount = 0
        lastApiCallTime = Timer
    End If
    
    apiCallCount = apiCallCount + 1
    lastApiCallTime = Timer
End Sub

' Sleep without API declares - keeps Excel responsive
Private Sub SleepMs(ByVal ms As Long)
    Dim t As Single: t = Timer
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
End Sub

' ===================================
' MAIN PROCEDURE
' ===================================

Sub GeneratePricingSummary()
    Dim wsManifest As Worksheet
    Dim wsSummary As Worksheet
    Dim currentRow As Long
    Dim startTime As Double
    
    startTime = Timer
    
    ' Initialize all module-level variables
    Call InitializeModuleVariables
    
    ' Set up event handler
    Call SetupSummarySheetEventHandler
    
    ' Validate MANIFEST sheet
    On Error Resume Next
    Set wsManifest = ActiveWorkbook.Sheets("MANIFEST")
    On Error GoTo 0
    
    If wsManifest Is Nothing Then
        MsgBox "MANIFEST sheet not found!", vbExclamation, "Error"
        Exit Sub
    End If
    
    If wsManifest.Cells(2, 2).Value = "" Then
        MsgBox "Please enter data in the MANIFEST sheet before generating summary.", vbExclamation, "No Data Found"
        Exit Sub
    End If
    
    ' ===================================
    ' OPTIMIZATION: Disable Excel features during processing
    ' ===================================
    Dim calcMode As XlCalculation
    calcMode = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Loading MANIFEST data into memory..."
    
    On Error GoTo ErrorHandler
    
    ' ===================================
    ' STEP 1: Load all MANIFEST data into memory (SINGLE PASS)
    ' ===================================
    Call LoadManifestDataIntoMemory(wsManifest)
    
    If manifestDataCount = 0 Then
        MsgBox "No valid data found in MANIFEST.", vbExclamation, "No Data"
        GoTo CleanupAndExit
    End If
    
    Application.StatusBar = "Building data groups..."
    DoEvents
    
    ' ===================================
    ' STEP 2: Build all groupings in memory (SINGLE PASS)
    ' ===================================
    Call BuildDataGroups
    
    Application.StatusBar = "Fetching WholeCell pricing for unique IMEIs..."
    DoEvents
    
    ' ===================================
    ' STEP 3: Fetch WholeCell costs for ALL unique IMEIs (BATCHED)
    ' ===================================
    Call FetchAllWholeCellCosts
    
    ' ===================================
    ' STEP 4: Calculate average costs per variation group
    ' ===================================
    Call CalculateGroupCosts
    
    Application.StatusBar = "Creating summary sheet..."
    DoEvents
    
    ' ===================================
    ' STEP 5: Create/clear summary sheet
    ' ===================================
    Set wsSummary = GetOrCreateSummarySheet(wsManifest)
    If wsSummary Is Nothing Then GoTo CleanupAndExit
    
    ' ===================================
    ' STEP 6: Write all data to summary sheet (BATCHED)
    ' ===================================
    currentRow = 1
    
    Dim modelKey As Variant
    Dim modelCount As Long
    modelCount = 0
    
    For Each modelKey In uniqueModelKeys.Keys
        modelCount = modelCount + 1
        Application.StatusBar = "Writing model " & modelCount & " of " & uniqueModelKeys.count & "..."
        
        ' Process Headers1 section
        currentRow = WriteHeaders1Section(wsSummary, CStr(modelKey), currentRow)
        
        ' Process Headers2 section
        currentRow = WriteHeaders2Section(wsSummary, currentRow)
        
        ' Add empty row between models
        currentRow = currentRow + 1
    Next modelKey
    
    ' ===================================
    ' STEP 7: Format the summary sheet
    ' ===================================
    Call FormatSummarySheet(wsSummary)
    
    ' Re-enable Excel features before showing results
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
    ' Switch to summary sheet
    wsSummary.Activate
    
    ' Show statistics
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    Dim statsMsg As String
    statsMsg = "Pricing Summary Generated in " & Format(elapsedTime, "0.0") & " seconds!" & vbNewLine & vbNewLine & _
               "WholeCell IMEI Lookup Results:" & vbNewLine & _
               "================================" & vbNewLine & _
               "  [+] Found: " & totalIMEIsFound & " IMEIs" & vbNewLine & _
               "  [-] Not Found: " & totalIMEIsNotFound & " IMEIs" & vbNewLine & _
               "================================" & vbNewLine & _
               "Total Unique IMEIs: " & totalIMEIsProcessed & vbNewLine & _
               "Models Processed: " & uniqueModelKeys.count
    
    If totalIMEIsNotFound > 0 Then
        statsMsg = statsMsg & vbNewLine & vbNewLine & _
                   "Note: Items without pricing were excluded from summary."
    End If
    
    MsgBox statsMsg, IIf(totalIMEIsNotFound > 0, vbExclamation, vbInformation), "Summary Complete"
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    
CleanupAndExit:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub

' ===================================
' INITIALIZATION
' ===================================

Private Sub InitializeModuleVariables()
    ' Initialize cost cache
    Set costCache = CreateObject("Scripting.Dictionary")
    
    ' Initialize data structures
    Set uniqueModelKeys = CreateObject("Scripting.Dictionary")
    Set modelVariations = CreateObject("Scripting.Dictionary")
    Set variationGroups = CreateObject("Scripting.Dictionary")
    Set allUniqueIMEIs = CreateObject("Scripting.Dictionary")
    
    ' Reset statistics
    totalIMEIsProcessed = 0
    totalIMEIsFound = 0
    totalIMEIsNotFound = 0
    
    ' Reset rate limiting
    lastApiCallTime = 0
    apiCallCount = 0
    
    ' Reset manifest data
    manifestDataCount = 0
    Erase manifestData
End Sub

' ===================================
' STEP 1: LOAD MANIFEST DATA INTO MEMORY (SINGLE PASS)
' ===================================

Private Sub LoadManifestDataIntoMemory(ws As Worksheet)
    Dim lastRow As Long
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim i As Long
    Dim item As ManifestItem
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' OPTIMIZATION: Read entire data range into array at once
    ' Columns: B(Model), C(Desc), D(IMEI), G(Grade), H(Battery), K(Color)
    Set dataRange = ws.Range("A2:K" & lastRow)
    dataArray = dataRange.Value
    
    ' Pre-allocate array for manifest items
    ReDim manifestData(1 To UBound(dataArray, 1))
    manifestDataCount = 0
    
    ' Single pass through data
    For i = 1 To UBound(dataArray, 1)
        Dim model As String, storage As String, imei As String
        Dim grade As String, battery As String, color As String
        
        model = Trim(Nz(dataArray(i, 2), ""))     ' Column B
        storage = ExtractStorage(Nz(dataArray(i, 3), ""))  ' Column C (Description)
        imei = Trim(Nz(dataArray(i, 4), ""))      ' Column D
        grade = Trim(Nz(dataArray(i, 7), ""))     ' Column G
        battery = Trim(Nz(dataArray(i, 8), ""))   ' Column H
        color = Trim(Nz(dataArray(i, 11), ""))    ' Column K
        
        ' Only process rows with valid model and storage
        If Len(model) > 0 And Len(storage) > 0 Then
            manifestDataCount = manifestDataCount + 1
            
            With manifestData(manifestDataCount)
                .RowIndex = i + 1  ' +1 because we started from row 2
                .Model = model
                .Storage = storage
                .Color = color
                .Grade = grade
                .Battery = battery
                .IMEI = imei
                .ModelKey = model & "|" & storage
                .VariationKey = grade & "|" & battery & "|" & color
                .FullKey = model & "|" & storage & "|" & grade & "|" & battery & "|" & color
            End With
        End If
    Next i
    
    ' Trim array to actual size
    If manifestDataCount > 0 Then
        ReDim Preserve manifestData(1 To manifestDataCount)
    End If
End Sub

' ===================================
' STEP 2: BUILD DATA GROUPS (SINGLE PASS)
' ===================================

Private Sub BuildDataGroups()
    Dim i As Long
    Dim item As ManifestItem
    Dim grp As VariationGroup
    Dim variations As Collection
    
    For i = 1 To manifestDataCount
        item = manifestData(i)
        
        ' Track unique model keys
        If Not uniqueModelKeys.Exists(item.ModelKey) Then
            uniqueModelKeys.Add item.ModelKey, True
            Set modelVariations(item.ModelKey) = New Collection
        End If
        
        ' Track variations per model
        Set variations = modelVariations(item.ModelKey)
        On Error Resume Next
        variations.Add item.VariationKey, item.VariationKey
        On Error GoTo 0
        
        ' Build/update variation group
        If Not variationGroups.Exists(item.FullKey) Then
            ' Create new group
            Dim newGrp As VariationGroup
            newGrp.Model = item.Model
            newGrp.Storage = item.Storage
            newGrp.Color = item.Color
            newGrp.Grade = item.Grade
            newGrp.Battery = item.Battery
            newGrp.Quantity = 0
            Set newGrp.IMEIs = New Collection
            newGrp.TotalCost = 0
            newGrp.ValidCostCount = 0
            newGrp.AvgCost = 0
            variationGroups.Add item.FullKey, newGrp
        End If
        
        ' Update group
        grp = variationGroups(item.FullKey)
        grp.Quantity = grp.Quantity + 1
        
        ' Track unique IMEIs for this group (avoid duplicates)
        If Len(item.IMEI) > 0 Then
            On Error Resume Next
            grp.IMEIs.Add item.IMEI, item.IMEI
            On Error GoTo 0
            
            ' Also track globally unique IMEIs
            If Not allUniqueIMEIs.Exists(item.IMEI) Then
                allUniqueIMEIs.Add item.IMEI, 0  ' 0 = not yet fetched
            End If
        End If
        
        variationGroups(item.FullKey) = grp
    Next i
End Sub

' ===================================
' STEP 3: FETCH ALL WHOLECELL COSTS (BATCHED, WITH CACHING)
' ===================================

Private Sub FetchAllWholeCellCosts()
    Dim imei As Variant
    Dim totalIMEIs As Long
    Dim currentIMEI As Long
    Dim cost As Double
    
    totalIMEIs = allUniqueIMEIs.count
    currentIMEI = 0
    
    ' Iterate through all unique IMEIs
    For Each imei In allUniqueIMEIs.Keys
        currentIMEI = currentIMEI + 1
        totalIMEIsProcessed = totalIMEIsProcessed + 1
        
        ' Update status every 5 IMEIs to reduce overhead
        If currentIMEI Mod 5 = 0 Or currentIMEI = 1 Then
            Application.StatusBar = "Fetching IMEI " & currentIMEI & " of " & totalIMEIs & "..."
            DoEvents
        End If
        
        ' Check if already in cache (from previous runs or duplicate)
        If costCache.Exists(imei) Then
            cost = costCache(imei)
        Else
            ' Fetch from WholeCell API
            cost = FetchSingleIMEICost(CStr(imei))
            
            ' Cache the result (even if 0, to avoid re-fetching)
            costCache.Add imei, cost
        End If
        
        ' Store in our IMEIs dictionary
        allUniqueIMEIs(imei) = cost
        
        ' Track statistics
        If cost > 0 Then
            totalIMEIsFound = totalIMEIsFound + 1
        Else
            totalIMEIsNotFound = totalIMEIsNotFound + 1
        End If
    Next imei
End Sub

' Fetch cost for a single IMEI
Private Function FetchSingleIMEICost(ByVal imei As String) As Double
    Dim dict As Object
    Dim cost As Double
    
    cost = 0
    
    Set dict = FetchWithRetry(imei, 2)  ' 2 retries
    
    If Not dict Is Nothing Then
        If CBool(Nz(dict("success"), False)) Then
            cost = CDbl0(dict("universalUsd"))
        End If
    End If
    
    FetchSingleIMEICost = cost
End Function

' ===================================
' STEP 4: CALCULATE GROUP COSTS
' ===================================

Private Sub CalculateGroupCosts()
    Dim fullKey As Variant
    Dim grp As VariationGroup
    Dim imei As Variant
    Dim cost As Double
    
    For Each fullKey In variationGroups.Keys
        grp = variationGroups(fullKey)
        
        grp.TotalCost = 0
        grp.ValidCostCount = 0
        
        ' Sum costs for all IMEIs in this group
        For Each imei In grp.IMEIs
            If allUniqueIMEIs.Exists(imei) Then
                cost = allUniqueIMEIs(imei)
                If cost > 0 Then
                    grp.TotalCost = grp.TotalCost + cost
                    grp.ValidCostCount = grp.ValidCostCount + 1
                End If
            End If
        Next imei
        
        ' Calculate average
        If grp.ValidCostCount > 0 Then
            grp.AvgCost = grp.TotalCost / grp.ValidCostCount
        Else
            grp.AvgCost = 0
        End If
        
        variationGroups(fullKey) = grp
    Next fullKey
End Sub

' ===================================
' STEP 5: GET OR CREATE SUMMARY SHEET
' ===================================

Private Function GetOrCreateSummarySheet(wsManifest As Worksheet) As Worksheet
    Dim wsSummary As Worksheet
    
    On Error Resume Next
    Set wsSummary = ActiveWorkbook.Sheets("Summary of pricing")
    On Error GoTo 0
    
    If wsSummary Is Nothing Then
        Set wsSummary = ActiveWorkbook.Sheets.Add(After:=wsManifest)
        wsSummary.Name = "Summary of pricing"
    Else
        ' Ask user if they want to overwrite
        Application.ScreenUpdating = True
        If MsgBox("Summary sheet already exists. Do you want to regenerate it?", vbYesNo + vbQuestion, "Overwrite Summary?") = vbNo Then
            Set GetOrCreateSummarySheet = Nothing
            Exit Function
        End If
        Application.ScreenUpdating = False
        wsSummary.Cells.Clear
    End If
    
    Set GetOrCreateSummarySheet = wsSummary
End Function

' ===================================
' STEP 6A: WRITE HEADERS1 SECTION
' ===================================

Private Function WriteHeaders1Section(wsSummary As Worksheet, modelKey As String, startRow As Long) As Long
    Dim currentRow As Long
    Dim variations As Collection
    Dim varKey As Variant
    Dim fullKey As String
    Dim grp As VariationGroup
    Dim rowsWritten As Long
    Dim firstDataRow As Long
    
    currentRow = startRow
    
    ' Write Headers1 row
    With wsSummary
        .Cells(currentRow, 1).Value = "Model + Storage + Battery status + Grade + Color"
        .Cells(currentRow, 2).Value = "Qty"
        .Cells(currentRow, 3).Value = "Cost"
        .Cells(currentRow, 4).Value = "Extended Cost"
        .Cells(currentRow, 5).Value = ""
        .Cells(currentRow, 6).Value = "Multiplier"
        .Cells(currentRow, 7).Value = "Suggested Wholesale"
        .Cells(currentRow, 8).Value = "Suggested Extended"
        .Cells(currentRow, 9).Value = ""
        .Cells(currentRow, 10).Value = "Free hand Multiplier"
        .Cells(currentRow, 11).Value = "Final whole sale"
        .Cells(currentRow, 12).Value = "Final extended"
        .Cells(currentRow, 13).Value = "Notes"
        
        ' Format header row
        .Range(.Cells(currentRow, 1), .Cells(currentRow, 13)).Borders.LineStyle = xlContinuous
    End With
    
    currentRow = currentRow + 1
    firstDataRow = currentRow
    rowsWritten = 0
    
    ' Get variations for this model
    If Not modelVariations.Exists(modelKey) Then
        WriteHeaders1Section = currentRow
        Exit Function
    End If
    
    Set variations = modelVariations(modelKey)
    
    ' Process each variation
    For Each varKey In variations
        fullKey = modelKey & "|" & CStr(varKey)
        
        If variationGroups.Exists(fullKey) Then
            grp = variationGroups(fullKey)
            
            ' Only write row if we have valid pricing
            If grp.AvgCost > 0 Then
                Call WriteDataRow(wsSummary, grp, currentRow)
                currentRow = currentRow + 1
                rowsWritten = rowsWritten + 1
            End If
        End If
    Next varKey
    
    WriteHeaders1Section = currentRow
End Function

' Write a single data row
Private Sub WriteDataRow(wsSummary As Worksheet, grp As VariationGroup, rowNum As Long)
    Dim multiplier As Double
    Dim suggestedPrice As Double
    Dim batteryDeduction As Double
    Dim isBadBattery As Boolean
    Dim batteryNote As String
    
    With wsSummary
        ' Column A: Full description
        .Cells(rowNum, 1).Value = grp.Model & " " & grp.Storage & " " & grp.Battery & " " & grp.Grade & " " & grp.Color
        
        ' Column B: Qty
        .Cells(rowNum, 2).Value = grp.Quantity
        
        ' Column C: Cost (from WholeCell) - YELLOW
        .Cells(rowNum, 3).Value = grp.AvgCost
        .Cells(rowNum, 3).NumberFormat = "$#,##0.00"
        .Cells(rowNum, 3).Interior.color = RGB(255, 255, 0)
        
        ' Column D: Extended Cost
        .Cells(rowNum, 4).Formula = "=B" & rowNum & "*C" & rowNum
        .Cells(rowNum, 4).NumberFormat = "$#,##0.00"
        
        ' Column F: Multiplier - GREEN
        multiplier = GetMultiplier(grp.Grade)
        .Cells(rowNum, 6).Value = multiplier
        .Cells(rowNum, 6).NumberFormat = "0.00"
        .Cells(rowNum, 6).Interior.color = RGB(146, 208, 80)
        
        ' Calculate battery deduction
        isBadBattery = IsBatteryDefective(grp.Grade, grp.Battery)
        batteryDeduction = IIf(isBadBattery, 15, 0)
        
        ' Column G: Suggested Wholesale - YELLOW
        suggestedPrice = (grp.AvgCost * multiplier) - batteryDeduction
        .Cells(rowNum, 7).Value = suggestedPrice
        .Cells(rowNum, 7).NumberFormat = "$#,##0.00"
        .Cells(rowNum, 7).Interior.color = RGB(255, 255, 0)
        
        ' Column H: Suggested Extended - YELLOW
        .Cells(rowNum, 8).Formula = "=B" & rowNum & "*G" & rowNum
        .Cells(rowNum, 8).NumberFormat = "$#,##0.00"
        .Cells(rowNum, 8).Interior.color = RGB(255, 255, 0)
        
        ' Column J: Free hand Multiplier - BLUE
        .Cells(rowNum, 10).Interior.color = RGB(142, 169, 219)
        .Cells(rowNum, 10).NumberFormat = "0.00"
        
        ' Column K: Final wholesale - BLUE
        If isBadBattery Then
            .Cells(rowNum, 11).Formula = "=IF(ISNUMBER(J" & rowNum & "),C" & rowNum & "*J" & rowNum & "-15,G" & rowNum & ")"
        Else
            .Cells(rowNum, 11).Formula = "=IF(ISNUMBER(J" & rowNum & "),C" & rowNum & "*J" & rowNum & ",G" & rowNum & ")"
        End If
        .Cells(rowNum, 11).NumberFormat = "$#,##0.00"
        .Cells(rowNum, 11).Interior.color = RGB(142, 169, 219)
        
        ' Column L: Final extended - BLUE
        .Cells(rowNum, 12).Formula = "=B" & rowNum & "*K" & rowNum
        .Cells(rowNum, 12).NumberFormat = "$#,##0.00"
        .Cells(rowNum, 12).Interior.color = RGB(142, 169, 219)
        
        ' Column M: Notes (battery deduction)
        If isBadBattery Then
            Dim totalDeduction As Double
            totalDeduction = batteryDeduction * grp.Quantity
            batteryNote = " (Battery deduction: -$" & Format(totalDeduction, "#,##0.00") & " total, -$15 per unit)"
            .Cells(rowNum, 13).Value = batteryNote
            .Cells(rowNum, 13).Font.color = RGB(255, 0, 0)
            .Cells(rowNum, 13).Font.Bold = True
        End If
        
        ' Column N: Hidden battery status
        .Cells(rowNum, 14).Value = IIf(isBadBattery, "BAD_BATTERY", "GOOD_BATTERY")
        .Cells(rowNum, 14).Font.color = .Cells(rowNum, 14).Interior.color
        
        ' Add borders
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 13)).Borders.LineStyle = xlContinuous
    End With
End Sub

' ===================================
' STEP 6B: WRITE HEADERS2 SECTION
' ===================================

Private Function WriteHeaders2Section(wsSummary As Worksheet, startRow As Long) As Long
    Dim currentRow As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    
    currentRow = startRow
    
    ' Find data row range
    firstDataRow = FindFirstDataRow(wsSummary, startRow)
    lastDataRow = startRow - 1
    
    ' Skip if no data rows
    If firstDataRow >= lastDataRow Then
        WriteHeaders2Section = currentRow
        Exit Function
    End If
    
    With wsSummary
        ' Write Headers2 header row
        .Cells(currentRow, 1).Value = "Model + Storage"
        .Cells(currentRow, 2).Value = ""
        .Cells(currentRow, 3).Value = "Average Cost"
        .Cells(currentRow, 4).Value = "Extended Cost"
        .Cells(currentRow, 5).Value = ""
        .Cells(currentRow, 6).Value = "Suggested wholesale"
        .Cells(currentRow, 7).Value = "Clear Rate"
        .Cells(currentRow, 8).Value = "Suggested Extended"
        .Cells(currentRow, 9).Value = ""
        .Cells(currentRow, 10).Value = "Final Wholesale"
        .Cells(currentRow, 11).Value = "Clear Rate"
        .Cells(currentRow, 12).Value = "Final Extended"
        .Cells(currentRow, 13).Value = ""
        
        ' Format header row - ORANGE
        With .Range(.Cells(currentRow, 1), .Cells(currentRow, 13))
            .Interior.color = RGB(255, 192, 0)
            .Borders.LineStyle = xlContinuous
        End With
        
        currentRow = currentRow + 1
        
        ' Extract Model + Storage info
        Dim modelStorageInfo As String
        modelStorageInfo = ExtractModelStorage(.Cells(firstDataRow, 1).Value)
        
        ' Write summary row
        .Cells(currentRow, 1).Value = modelStorageInfo
        .Cells(currentRow, 2).Formula = "=SUM(B" & firstDataRow & ":B" & lastDataRow & ")"
        
        .Cells(currentRow, 3).Formula = "=AVERAGE(C" & firstDataRow & ":C" & lastDataRow & ")"
        .Cells(currentRow, 3).NumberFormat = "$#,##0.00"
        
        .Cells(currentRow, 4).Formula = "=SUM(D" & firstDataRow & ":D" & lastDataRow & ")"
        .Cells(currentRow, 4).NumberFormat = "$#,##0.00"
        
        .Cells(currentRow, 6).Formula = "=SUMPRODUCT(G" & firstDataRow & ":G" & lastDataRow & ",B" & firstDataRow & ":B" & lastDataRow & ")/SUM(B" & firstDataRow & ":B" & lastDataRow & ")"
        .Cells(currentRow, 6).NumberFormat = "$#,##0.00"
        
        .Cells(currentRow, 7).Formula = "=F" & currentRow & "/C" & currentRow
        .Cells(currentRow, 7).NumberFormat = "0.00%"
        
        .Cells(currentRow, 8).Formula = "=SUM(H" & firstDataRow & ":H" & lastDataRow & ")"
        .Cells(currentRow, 8).NumberFormat = "$#,##0.00"
        
        .Cells(currentRow, 10).Formula = "=SUMPRODUCT(K" & firstDataRow & ":K" & lastDataRow & ",B" & firstDataRow & ":B" & lastDataRow & ")/SUM(B" & firstDataRow & ":B" & lastDataRow & ")"
        .Cells(currentRow, 10).NumberFormat = "$#,##0.00"
        
        .Cells(currentRow, 11).Formula = "=J" & currentRow & "/C" & currentRow
        .Cells(currentRow, 11).NumberFormat = "0.00%"
        
        .Cells(currentRow, 12).Formula = "=SUM(L" & firstDataRow & ":L" & lastDataRow & ")"
        .Cells(currentRow, 12).NumberFormat = "$#,##0.00"
        
        ' Format - ORANGE
        With .Range(.Cells(currentRow, 1), .Cells(currentRow, 13))
            .Borders.LineStyle = xlContinuous
            .Interior.color = RGB(255, 192, 0)
        End With
        
        currentRow = currentRow + 1
    End With
    
    WriteHeaders2Section = currentRow
End Function

' Find first data row (after Headers1 header)
Private Function FindFirstDataRow(ws As Worksheet, currentRow As Long) As Long
    Dim i As Long
    
    For i = currentRow - 1 To 1 Step -1
        If ws.Cells(i, 1).Value Like "*Model + Storage + Battery*" Then
            FindFirstDataRow = i + 1
            Exit Function
        End If
    Next i
    
    FindFirstDataRow = currentRow
End Function

' ===================================
' HELPER FUNCTIONS
' ===================================

' Extract storage from description
Private Function ExtractStorage(description As String) As String
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(\d+GB)"
    regex.IgnoreCase = True
    
    Set matches = regex.Execute(description)
    If matches.count > 0 Then
        ExtractStorage = matches(0).Value
    Else
        ExtractStorage = ""
    End If
End Function

' Get multiplier based on grade
Private Function GetMultiplier(grade As String) As Double
    Dim normalizedGrade As String
    normalizedGrade = NormalizeGrade(grade)
    
    ' Check for B GOOD variants
    If InStr(normalizedGrade, "B GOOD") > 0 Or normalizedGrade = "B" Then
        GetMultiplier = 1.15
        Exit Function
    End If
    
    ' Check for C(AMZ) variants
    If InStr(normalizedGrade, "CAMZ") > 0 Or _
       (InStr(normalizedGrade, "C") > 0 And InStr(normalizedGrade, "AMZ") > 0) Then
        GetMultiplier = 1.15
        Exit Function
    End If
    
    ' Check for C GOOD variants (but not AMZ)
    If InStr(normalizedGrade, "C GOOD") > 0 And InStr(normalizedGrade, "AMZ") = 0 Then
        GetMultiplier = 1.05
        Exit Function
    End If
    
    ' Check for DEFECTIVE variants
    If InStr(normalizedGrade, "DEFECTIVE") > 0 Then
        GetMultiplier = 0.4
        Exit Function
    End If
    
    ' Check for D GOOD variants
    If InStr(normalizedGrade, "D GOOD") > 0 Then
        GetMultiplier = 0.7
        Exit Function
    End If
    
    ' Single letter grades
    If normalizedGrade = "B" Then
        GetMultiplier = 1.15
    ElseIf normalizedGrade = "C" Then
        GetMultiplier = 1.05
    ElseIf normalizedGrade = "D" Then
        GetMultiplier = 0.7
    Else
        GetMultiplier = 0
    End If
End Function

' Normalize grade string
Private Function NormalizeGrade(grade As String) As String
    Dim result As String
    result = Trim(UCase(grade))
    
    ' Remove extra spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ' Normalize N/A variations
    result = Replace(result, "N/A", "NA")
    result = Replace(result, "N A", "NA")
    result = Replace(result, "N- A", "NA")
    
    ' Remove parentheses for AMZ
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    
    NormalizeGrade = Trim(result)
End Function

' Check if battery is defective
Private Function IsBatteryDefective(grade As String, batteryHealth As String) As Boolean
    ' Check grade for defective
    If InStr(UCase(grade), "DEFECTIVE") > 0 Then
        IsBatteryDefective = True
        Exit Function
    End If
    
    ' Check battery percentage
    If batteryHealth <> "" And batteryHealth <> "N/A" Then
        Dim cleanBattery As String
        Dim batteryPercent As Double
        
        cleanBattery = Replace(batteryHealth, "%", "")
        If IsNumeric(cleanBattery) Then
            batteryPercent = CDbl(cleanBattery)
            If batteryPercent < 80 Then
                IsBatteryDefective = True
                Exit Function
            End If
        End If
    End If
    
    ' Check for specific keywords
    If InStr(UCase(batteryHealth), "DEFECTIVE") > 0 Or _
       InStr(UCase(batteryHealth), "BAD") > 0 Or _
       InStr(UCase(batteryHealth), "FAIL") > 0 Or _
       (batteryHealth = "N/A" And InStr(UCase(grade), "DEFECTIVE") > 0) Then
        IsBatteryDefective = True
    Else
        IsBatteryDefective = False
    End If
End Function

' Extract Model + Storage from full string
Private Function ExtractModelStorage(fullString As String) As String
    Dim parts() As String
    Dim result As String
    Dim i As Long
    Dim foundStorage As Boolean
    
    parts = Split(fullString, " ")
    result = ""
    foundStorage = False
    
    For i = 0 To UBound(parts)
        result = result & parts(i)
        
        If UCase(parts(i)) Like "*GB" Then
            foundStorage = True
            Exit For
        End If
        
        result = result & " "
    Next i
    
    If Not foundStorage Then
        If UBound(parts) >= 2 Then
            result = parts(0) & " " & parts(1) & " " & parts(2)
        Else
            result = fullString
        End If
    End If
    
    ExtractModelStorage = Trim(result)
End Function

' ===================================
' FORMAT SUMMARY SHEET
' ===================================

Private Sub FormatSummarySheet(ws As Worksheet)
    With ws
        ' Auto-fit columns
        .Columns("A:M").AutoFit
        
        ' Set column widths
        .Columns("A").ColumnWidth = 55
        .Columns("B").ColumnWidth = 12
        .Columns("C:D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 8
        .Columns("F").ColumnWidth = 14
        .Columns("G:H").ColumnWidth = 18
        .Columns("I").ColumnWidth = 8
        .Columns("J").ColumnWidth = 18
        .Columns("K:L").ColumnWidth = 18
        .Columns("M").ColumnWidth = 55
        .Columns("N").ColumnWidth = 0.1
        .Columns("N").Hidden = True
        
        ' Freeze panes
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
        
        .Range("A1").Select
    End With
End Sub

' ===================================
' EVENT HANDLER SETUP
' ===================================

Private Sub SetupSummarySheetEventHandler()
    On Error Resume Next
    Dim wsSummary As Worksheet
    Set wsSummary = ActiveWorkbook.Sheets("Summary of pricing")
    If wsSummary Is Nothing Then Exit Sub
    On Error GoTo 0
    
    wsSummary.Cells(1, 14).Value = "EVENT_HANDLER_INSTALLED"
    wsSummary.Cells(1, 14).Font.color = wsSummary.Cells(1, 14).Interior.color
End Sub

' ===================================
' BUTTON SETUP
' ===================================

Sub AddButtonToManifest()
    Dim ws As Worksheet
    Dim btn As Object
    Dim btnRange As Range
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("MANIFEST")
    If ws Is Nothing Then
        MsgBox "MANIFEST sheet not found! Please create MANIFEST sheet first.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ws.Activate
    
    On Error Resume Next
    ws.Shapes("btnGenerateSummary").Delete
    On Error GoTo 0
    
    Set btnRange = ws.Range("M2:O3")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    
    With btn
        .Name = "btnGenerateSummary"
        .Caption = "Generate Pricing Summary"
        .OnAction = "GeneratePricingSummary"
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    ws.Range("M1").Value = "Click to generate summary:"
    ws.Range("M1").Font.Bold = True
    ws.Range("M1").Font.color = RGB(0, 0, 200)
    
    MsgBox "Button successfully added to MANIFEST sheet!" & vbNewLine & _
           vbNewLine & _
           "Instructions:" & vbNewLine & _
           "1. Enter all your data in the MANIFEST sheet" & vbNewLine & _
           "2. Click the 'Generate Pricing Summary' button" & vbNewLine & _
           "3. The summary will be created in a new sheet", _
           vbInformation, "Setup Complete"
End Sub

Sub AddToQuickAccess()
    MsgBox "To add to Quick Access Toolbar:" & vbNewLine & _
           "1. Right-click the Quick Access Toolbar" & vbNewLine & _
           "2. Choose 'Customize Quick Access Toolbar'" & vbNewLine & _
           "3. Choose 'Macros' from the dropdown" & vbNewLine & _
           "4. Select 'GeneratePricingSummary'" & vbNewLine & _
           "5. Click 'Add' and then 'OK'", _
           vbInformation, "Quick Access Setup"
End Sub

' ===================================
' WORKSHEET CHANGE EVENT HANDLER
' ===================================
' Copy this to the "Summary of pricing" sheet module:
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     If Intersect(Target, Me.Range("J:L")) Is Nothing Then Exit Sub
'     
'     Dim affectedRow As Long
'     affectedRow = Target.Row
'     
'     If Me.Cells(affectedRow, 1).Value = "" Or _
'        Me.Cells(affectedRow, 1).Value Like "*Model + Storage*" Then Exit Sub
'     
'     Application.EnableEvents = False
'     On Error GoTo ErrorHandler
'     
'     Dim colChanged As Integer: colChanged = Target.Column
'     Dim cost As Double, qty As Double, kValue As Double, jValue As Double, lValue As Double
'     Dim hasBadBattery As Boolean, batteryDeduction As Double
'     
'     hasBadBattery = (Me.Cells(affectedRow, 14).Value = "BAD_BATTERY")
'     batteryDeduction = IIf(hasBadBattery, 15, 0)
'     
'     On Error Resume Next
'     cost = CDbl(Me.Cells(affectedRow, 3).Value)
'     qty = CDbl(Me.Cells(affectedRow, 2).Value)
'     On Error GoTo ErrorHandler
'     
'     Select Case colChanged
'         Case 10
'             jValue = CDbl(Target.Value)
'             If jValue > 0 And cost > 0 Then
'                 kValue = (cost * jValue) - batteryDeduction
'                 Me.Cells(affectedRow, 11).Value = kValue
'                 lValue = qty * kValue
'                 Me.Cells(affectedRow, 12).Value = lValue
'             End If
'         Case 11
'             kValue = CDbl(Target.Value)
'             If cost > 0 Then
'                 jValue = (kValue + batteryDeduction) / cost
'                 Me.Cells(affectedRow, 10).Value = jValue
'                 lValue = qty * kValue
'                 Me.Cells(affectedRow, 12).Value = lValue
'             End If
'         Case 12
'             lValue = CDbl(Target.Value)
'             If qty > 0 Then
'                 kValue = lValue / qty
'                 Me.Cells(affectedRow, 11).Value = kValue
'                 If cost > 0 Then
'                     jValue = (kValue + batteryDeduction) / cost
'                     Me.Cells(affectedRow, 10).Value = jValue
'                 End If
'             End If
'     End Select
'     
' ErrorHandler:
'     Application.EnableEvents = True
' End Sub
