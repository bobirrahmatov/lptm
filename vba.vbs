Attribute VB_Name = "AggregatePortfolioData"
' =========================================
' Portfolio Analytics Data Aggregator
' =========================================
' This VBA script aggregates large portfolio data into a lightweight JSON format
' optimized for the portfolioAnalytics.html dashboard
'
' INSTRUCTIONS:
' 1. Open your Excel file with the portfolio data
' 2. Press ALT+F11 to open VBA Editor
' 3. Insert > Module
' 4. Paste this code
' 5. Update the constants below to match your data
' 6. Run the "AggregateAndExportToJSON" macro
' 7. Upload the generated JSON file to Confluence as an attachment

Option Explicit

' ===== CONFIGURATION =====
' Update these constants to match your Excel sheet and column names
Const SOURCE_SHEET_NAME As String = "Sheet1" ' Your sheet name
Const OUTPUT_FILE_NAME As String = "portfolioData.json" ' Output filename
Const MAX_DETAIL_RECORDS As Long = 1000 ' Max records to include in detail table (for performance)

' Column mappings - Update these to match your actual column headers
Const COL_REPORT_DATE As String = "Report_Date"
Const COL_REGION As String = "Region"
Const COL_FACILITY_ID As String = "Facility_ID"
Const COL_RELATIONSHIP_NAME As String = "Relationship_Name"
Const COL_RELATIONSHIP_ID As String = "Relationship_ID"
Const COL_PRODUCT_PROGRAM As String = "Product_Program"
Const COL_FAC_AMOUNT As String = "Fac_Amount"
Const COL_DIRECT_OS As String = "Direct_OS"
Const COL_MATURITY_DATE As String = "Maturity_Date"
Const COL_CA_EXPIRATION_DATE As String = "CA_Expiration_Date"
Const COL_CAID As String = "CAID"
Const COL_FACILITY_TYPE As String = "Facility_Type"
Const COL_MANAGEMENT_STATUS As String = "Management_Status"
Const COL_COMMITMENT As String = "Committed/Uncommitted"
Const COL_CONTROL_UNIT As String = "Control_Unit"
Const COL_TEAM_LEAD As String = "Underwriting_Team_Lead"
Const COL_UNDERWRITER As String = "Lead_Underwriter"

' ===== MAIN SUBROUTINE =====
Sub AggregateAndExportToJSON()
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Debug.Print "========================================="
    Debug.Print "Portfolio Data Aggregation Started"
    Debug.Print "========================================="
    
    ' Get source data
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SOURCE_SHEET_NAME)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Debug.Print "Total rows found: " & (lastRow - 1)
    
    ' Get column indices
    Dim headers As Range
    Set headers = ws.Rows(1)
    
    ' Build aggregated data structure
    Dim jsonData As String
    jsonData = "{"
    
    ' 1. Calculate Summary Metrics
    Debug.Print "Calculating summary metrics..."
    jsonData = jsonData & """metrics"": " & GetMetrics(ws, lastRow, headers) & ","
    
    ' 2. Aggregate Monthly Expiration Data
    Debug.Print "Aggregating monthly expirations..."
    jsonData = jsonData & """monthlyExpirations"": " & GetMonthlyExpirations(ws, lastRow, headers) & ","
    
    ' 3. Aggregate Regional Distribution
    Debug.Print "Aggregating regional distribution..."
    jsonData = jsonData & """regionalDistribution"": " & GetRegionalDistribution(ws, lastRow, headers) & ","
    
    ' 4. Get Top Relationships
    Debug.Print "Calculating top relationships..."
    jsonData = jsonData & """topRelationships"": " & GetTopRelationships(ws, lastRow, headers, 50) & ","
    
    ' 5. Get Filter Options
    Debug.Print "Extracting filter options..."
    jsonData = jsonData & """filterOptions"": " & GetFilterOptions(ws, lastRow, headers) & ","
    
    ' 6. Get Sample Detail Records (most recent)
    Debug.Print "Extracting sample detail records..."
    jsonData = jsonData & """detailRecords"": " & GetDetailRecords(ws, lastRow, headers, MAX_DETAIL_RECORDS)
    
    jsonData = jsonData & "}"
    
    ' Save to file
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & OUTPUT_FILE_NAME
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fileStream As Object
    Set fileStream = fso.CreateTextFile(filePath, True, False)
    fileStream.Write jsonData
    fileStream.Close
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Dim elapsed As Double
    elapsed = Timer - startTime
    
    Debug.Print "========================================="
    Debug.Print "SUCCESS! File saved to:"
    Debug.Print filePath
    Debug.Print "Time elapsed: " & Format(elapsed, "0.00") & " seconds"
    Debug.Print "========================================="
    
    MsgBox "Data aggregation complete!" & vbCrLf & vbCrLf & _
           "File saved to:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
           "Time: " & Format(elapsed, "0.00") & " seconds" & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Upload " & OUTPUT_FILE_NAME & " to your Confluence page" & vbCrLf & _
           "2. Update CONFLUENCE_FILE_NAME in portfolioAnalytics.html", _
           vbInformation, "Export Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    Debug.Print "ERROR: " & Err.Description
End Sub

' ===== HELPER FUNCTIONS =====

' Get column index by header name
Function GetColIndex(headers As Range, colName As String) As Long
    On Error Resume Next
    GetColIndex = Application.WorksheetFunction.Match(colName, headers, 0)
    If Err.Number <> 0 Then
        Debug.Print "Warning: Column '" & colName & "' not found"
        GetColIndex = 0
    End If
    On Error GoTo 0
End Function

' Calculate summary metrics
Function GetMetrics(ws As Worksheet, lastRow As Long, headers As Range) As String
    Dim colFacAmount As Long, colDirectOS As Long, colRelName As Long
    Dim colFacID As Long, colCAID As Long
    
    colFacAmount = GetColIndex(headers, COL_FAC_AMOUNT)
    colDirectOS = GetColIndex(headers, COL_DIRECT_OS)
    colRelName = GetColIndex(headers, COL_RELATIONSHIP_NAME)
    colFacID = GetColIndex(headers, COL_FACILITY_ID)
    colCAID = GetColIndex(headers, COL_CAID)
    
    Dim totalFacAmount As Double, totalDirectOS As Double
    Dim uniqueRels As Object, uniqueFacs As Object, uniqueCAs As Object
    Set uniqueRels = CreateObject("Scripting.Dictionary")
    Set uniqueFacs = CreateObject("Scripting.Dictionary")
    Set uniqueCAs = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lastRow
        ' Sum amounts
        If colFacAmount > 0 Then totalFacAmount = totalFacAmount + Val(ws.Cells(i, colFacAmount).Value)
        If colDirectOS > 0 Then totalDirectOS = totalDirectOS + Val(ws.Cells(i, colDirectOS).Value)
        
        ' Count uniques
        If colRelName > 0 And ws.Cells(i, colRelName).Value <> "" Then
            uniqueRels(ws.Cells(i, colRelName).Value) = 1
        End If
        If colFacID > 0 And ws.Cells(i, colFacID).Value <> "" Then
            uniqueFacs(ws.Cells(i, colFacID).Value) = 1
        End If
        If colCAID > 0 And ws.Cells(i, colCAID).Value <> "" Then
            uniqueCAs(ws.Cells(i, colCAID).Value) = 1
        End If
    Next i
    
    GetMetrics = "{" & _
        """totalRecords"": " & (lastRow - 1) & "," & _
        """totalFacilityAmount"": " & totalFacAmount & "," & _
        """totalDirectOS"": " & totalDirectOS & "," & _
        """uniqueRelationships"": " & uniqueRels.Count & "," & _
        """uniqueFacilities"": " & uniqueFacs.Count & "," & _
        """uniqueCAs"": " & uniqueCAs.Count & _
        "}"
End Function

' Aggregate monthly expirations
Function GetMonthlyExpirations(ws As Worksheet, lastRow As Long, headers As Range) As String
    Dim colMaturity As Long, colCAExp As Long
    colMaturity = GetColIndex(headers, COL_MATURITY_DATE)
    colCAExp = GetColIndex(headers, COL_CA_EXPIRATION_DATE)
    
    Dim facilityMonths As Object, caMonths As Object
    Set facilityMonths = CreateObject("Scripting.Dictionary")
    Set caMonths = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, monthKey As String, dateVal As Date
    
    ' Count facility maturities by month
    If colMaturity > 0 Then
        For i = 2 To lastRow
            If IsDate(ws.Cells(i, colMaturity).Value) Then
                dateVal = CDate(ws.Cells(i, colMaturity).Value)
                monthKey = Format(dateVal, "yyyy-mm")
                If facilityMonths.Exists(monthKey) Then
                    facilityMonths(monthKey) = facilityMonths(monthKey) + 1
                Else
                    facilityMonths(monthKey) = 1
                End If
            End If
        Next i
    End If
    
    ' Count CA expirations by month
    If colCAExp > 0 Then
        For i = 2 To lastRow
            If IsDate(ws.Cells(i, colCAExp).Value) Then
                dateVal = CDate(ws.Cells(i, colCAExp).Value)
                monthKey = Format(dateVal, "yyyy-mm")
                If caMonths.Exists(monthKey) Then
                    caMonths(monthKey) = caMonths(monthKey) + 1
                Else
                    caMonths(monthKey) = 1
                End If
            End If
        Next i
    End If
    
    ' Build JSON
    Dim json As String
    json = "{""facilities"": {"
    
    Dim keys() As Variant, j As Long
    If facilityMonths.Count > 0 Then
        keys = facilityMonths.Keys
        For j = 0 To UBound(keys)
            If j > 0 Then json = json & ","
            json = json & """" & keys(j) & """: " & facilityMonths(keys(j))
        Next j
    End If
    
    json = json & "}, ""cas"": {"
    
    If caMonths.Count > 0 Then
        keys = caMonths.Keys
        For j = 0 To UBound(keys)
            If j > 0 Then json = json & ","
            json = json & """" & keys(j) & """: " & caMonths(keys(j))
        Next j
    End If
    
    json = json & "}}"
    
    GetMonthlyExpirations = json
End Function

' Aggregate regional distribution
Function GetRegionalDistribution(ws As Worksheet, lastRow As Long, headers As Range) As String
    Dim colRegion As Long, colFacID As Long, colCAID As Long
    Dim colMaturity As Long, colCAExp As Long
    
    colRegion = GetColIndex(headers, COL_REGION)
    colFacID = GetColIndex(headers, COL_FACILITY_ID)
    colCAID = GetColIndex(headers, COL_CAID)
    colMaturity = GetColIndex(headers, COL_MATURITY_DATE)
    colCAExp = GetColIndex(headers, COL_CA_EXPIRATION_DATE)
    
    ' Track facilities and CAs expiring in next 90 days by region
    Dim regionFacilities As Object, regionCAs As Object
    Set regionFacilities = CreateObject("Scripting.Dictionary")
    Set regionCAs = CreateObject("Scripting.Dictionary")
    
    Dim today As Date, threeMonthsLater As Date
    today = Date
    threeMonthsLater = DateAdd("d", 90, today)
    
    Dim i As Long, region As String, dateVal As Date
    
    ' Count facilities by region (expiring soon)
    If colRegion > 0 And colMaturity > 0 And colFacID > 0 Then
        For i = 2 To lastRow
            If IsDate(ws.Cells(i, colMaturity).Value) Then
                dateVal = CDate(ws.Cells(i, colMaturity).Value)
                If dateVal >= today And dateVal <= threeMonthsLater Then
                    region = Trim(ws.Cells(i, colRegion).Value)
                    If region <> "" Then
                        If Not regionFacilities.Exists(region) Then
                            Set regionFacilities(region) = CreateObject("Scripting.Dictionary")
                        End If
                        regionFacilities(region)(ws.Cells(i, colFacID).Value) = 1
                    End If
                End If
            End If
        Next i
    End If
    
    ' Count CAs by region (expiring soon)
    If colRegion > 0 And colCAExp > 0 And colCAID > 0 Then
        For i = 2 To lastRow
            If IsDate(ws.Cells(i, colCAExp).Value) Then
                dateVal = CDate(ws.Cells(i, colCAExp).Value)
                If dateVal >= today And dateVal <= threeMonthsLater Then
                    region = Trim(ws.Cells(i, colRegion).Value)
                    If region <> "" Then
                        If Not regionCAs.Exists(region) Then
                            Set regionCAs(region) = CreateObject("Scripting.Dictionary")
                        End If
                        regionCAs(region)(ws.Cells(i, colCAID).Value) = 1
                    End If
                End If
            End If
        Next i
    End If
    
    ' Build JSON
    Dim json As String
    json = "{""facilityLevel"": {"
    
    Dim keys() As Variant, j As Long
    If regionFacilities.Count > 0 Then
        keys = regionFacilities.Keys
        For j = 0 To UBound(keys)
            If j > 0 Then json = json & ","
            json = json & """" & keys(j) & """: " & regionFacilities(keys(j)).Count
        Next j
    End If
    
    json = json & "}, ""caLevel"": {"
    
    If regionCAs.Count > 0 Then
        keys = regionCAs.Keys
        For j = 0 To UBound(keys)
            If j > 0 Then json = json & ","
            json = json & """" & keys(j) & """: " & regionCAs(keys(j)).Count
        Next j
    End If
    
    json = json & "}}"
    
    GetRegionalDistribution = json
End Function

' Get top relationships by amount
Function GetTopRelationships(ws As Worksheet, lastRow As Long, headers As Range, topN As Long) As String
    Dim colRelName As Long, colFacAmount As Long, colDirectOS As Long
    
    colRelName = GetColIndex(headers, COL_RELATIONSHIP_NAME)
    colFacAmount = GetColIndex(headers, COL_FAC_AMOUNT)
    colDirectOS = GetColIndex(headers, COL_DIRECT_OS)
    
    Dim relFacAmounts As Object, relDirectOS As Object
    Set relFacAmounts = CreateObject("Scripting.Dictionary")
    Set relDirectOS = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, relName As String
    
    ' Aggregate by relationship
    For i = 2 To lastRow
        If colRelName > 0 Then
            relName = Trim(ws.Cells(i, colRelName).Value)
            If relName <> "" Then
                ' Facility Amount
                If colFacAmount > 0 Then
                    If relFacAmounts.Exists(relName) Then
                        relFacAmounts(relName) = relFacAmounts(relName) + Val(ws.Cells(i, colFacAmount).Value)
                    Else
                        relFacAmounts(relName) = Val(ws.Cells(i, colFacAmount).Value)
                    End If
                End If
                
                ' Direct OS
                If colDirectOS > 0 Then
                    If relDirectOS.Exists(relName) Then
                        relDirectOS(relName) = relDirectOS(relName) + Val(ws.Cells(i, colDirectOS).Value)
                    Else
                        relDirectOS(relName) = Val(ws.Cells(i, colDirectOS).Value)
                    End If
                End If
            End If
        End If
    Next i
    
    ' Sort and get top N (simple bubble sort for small N)
    Dim sortedFac() As Variant, sortedDOS() As Variant
    sortedFac = GetTopNFromDict(relFacAmounts, topN)
    sortedDOS = GetTopNFromDict(relDirectOS, topN)
    
    ' Build JSON
    Dim json As String
    json = "{""byFacilityAmount"": ["
    
    Dim j As Long
    If Not IsEmpty(sortedFac) Then
        For j = 0 To UBound(sortedFac, 1)
            If j > 0 Then json = json & ","
            json = json & "{""name"": """ & EscapeJSON(sortedFac(j, 0)) & """, ""amount"": " & sortedFac(j, 1) & "}"
        Next j
    End If
    
    json = json & "], ""byDirectOS"": ["
    
    If Not IsEmpty(sortedDOS) Then
        For j = 0 To UBound(sortedDOS, 1)
            If j > 0 Then json = json & ","
            json = json & "{""name"": """ & EscapeJSON(sortedDOS(j, 0)) & """, ""amount"": " & sortedDOS(j, 1) & "}"
        Next j
    End If
    
    json = json & "]}"
    
    GetTopRelationships = json
End Function

' Get top N from dictionary
Function GetTopNFromDict(dict As Object, topN As Long) As Variant
    If dict.Count = 0 Then
        GetTopNFromDict = Empty
        Exit Function
    End If
    
    Dim keys() As Variant, items() As Variant
    keys = dict.Keys
    items = dict.Items
    
    Dim n As Long
    n = Application.Min(dict.Count, topN)
    
    ' Simple selection sort for top N
    Dim i As Long, j As Long, maxIdx As Long, tempKey As Variant, tempVal As Variant
    For i = 0 To n - 1
        maxIdx = i
        For j = i + 1 To dict.Count - 1
            If items(j) > items(maxIdx) Then maxIdx = j
        Next j
        
        If maxIdx <> i Then
            tempKey = keys(i): keys(i) = keys(maxIdx): keys(maxIdx) = tempKey
            tempVal = items(i): items(i) = items(maxIdx): items(maxIdx) = tempVal
        End If
    Next i
    
    ' Return top N
    Dim result() As Variant
    ReDim result(0 To n - 1, 0 To 1)
    For i = 0 To n - 1
        result(i, 0) = keys(i)
        result(i, 1) = items(i)
    Next i
    
    GetTopNFromDict = result
End Function

' Get filter options
Function GetFilterOptions(ws As Worksheet, lastRow As Long, headers As Range) As String
    Dim uniqueRegions As Object, uniqueProducts As Object, uniqueUnits As Object
    Dim uniqueFacTypes As Object, uniqueCommitment As Object, uniqueMgmt As Object
    Dim uniqueTeamLeads As Object, uniqueUnderwriters As Object
    
    Set uniqueRegions = CreateObject("Scripting.Dictionary")
    Set uniqueProducts = CreateObject("Scripting.Dictionary")
    Set uniqueUnits = CreateObject("Scripting.Dictionary")
    Set uniqueFacTypes = CreateObject("Scripting.Dictionary")
    Set uniqueCommitment = CreateObject("Scripting.Dictionary")
    Set uniqueMgmt = CreateObject("Scripting.Dictionary")
    Set uniqueTeamLeads = CreateObject("Scripting.Dictionary")
    Set uniqueUnderwriters = CreateObject("Scripting.Dictionary")
    
    Dim colRegion As Long, colProduct As Long, colUnit As Long, colFacType As Long
    Dim colCommitment As Long, colMgmt As Long, colTeamLead As Long, colUnderwriter As Long
    
    colRegion = GetColIndex(headers, COL_REGION)
    colProduct = GetColIndex(headers, COL_PRODUCT_PROGRAM)
    colUnit = GetColIndex(headers, COL_CONTROL_UNIT)
    colFacType = GetColIndex(headers, COL_FACILITY_TYPE)
    colCommitment = GetColIndex(headers, COL_COMMITMENT)
    colMgmt = GetColIndex(headers, COL_MANAGEMENT_STATUS)
    colTeamLead = GetColIndex(headers, COL_TEAM_LEAD)
    colUnderwriter = GetColIndex(headers, COL_UNDERWRITER)
    
    Dim i As Long, val As String
    
    For i = 2 To lastRow
        If colRegion > 0 Then
            val = Trim(ws.Cells(i, colRegion).Value)
            If val <> "" Then uniqueRegions(val) = 1
        End If
        If colProduct > 0 Then
            val = Trim(ws.Cells(i, colProduct).Value)
            If val <> "" Then uniqueProducts(val) = 1
        End If
        If colUnit > 0 Then
            val = Trim(ws.Cells(i, colUnit).Value)
            If val <> "" Then uniqueUnits(val) = 1
        End If
        If colFacType > 0 Then
            val = Trim(ws.Cells(i, colFacType).Value)
            If val <> "" Then uniqueFacTypes(val) = 1
        End If
        If colCommitment > 0 Then
            val = Trim(ws.Cells(i, colCommitment).Value)
            If val <> "" Then uniqueCommitment(val) = 1
        End If
        If colMgmt > 0 Then
            val = Trim(ws.Cells(i, colMgmt).Value)
            If val <> "" Then uniqueMgmt(val) = 1
        End If
        If colTeamLead > 0 Then
            val = Trim(ws.Cells(i, colTeamLead).Value)
            If val <> "" Then uniqueTeamLeads(val) = 1
        End If
        If colUnderwriter > 0 Then
            val = Trim(ws.Cells(i, colUnderwriter).Value)
            If val <> "" Then uniqueUnderwriters(val) = 1
        End If
    Next i
    
    ' Build JSON
    Dim json As String
    json = "{"
    json = json & """regions"": " & DictToJSONArray(uniqueRegions) & ","
    json = json & """productPrograms"": " & DictToJSONArray(uniqueProducts) & ","
    json = json & """controlUnits"": " & DictToJSONArray(uniqueUnits) & ","
    json = json & """facilityTypes"": " & DictToJSONArray(uniqueFacTypes) & ","
    json = json & """commitmentStatus"": " & DictToJSONArray(uniqueCommitment) & ","
    json = json & """managementStatus"": " & DictToJSONArray(uniqueMgmt) & ","
    json = json & """teamLeads"": " & DictToJSONArray(uniqueTeamLeads) & ","
    json = json & """underwriters"": " & DictToJSONArray(uniqueUnderwriters)
    json = json & "}"
    
    GetFilterOptions = json
End Function

' Get sample detail records (most recent)
Function GetDetailRecords(ws As Worksheet, lastRow As Long, headers As Range, maxRecords As Long) As String
    Dim json As String
    json = "["
    
    ' Determine columns to include
    Dim headerCount As Long
    headerCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Get last N records (or all if less than maxRecords)
    Dim startRow As Long
    startRow = Application.Max(2, lastRow - maxRecords + 1)
    
    Dim i As Long, j As Long, firstRecord As Boolean
    firstRecord = True
    
    For i = startRow To lastRow
        If Not firstRecord Then json = json & ","
        firstRecord = False
        
        json = json & "{"
        
        Dim firstCol As Boolean
        firstCol = True
        
        For j = 1 To headerCount
            Dim headerName As String
            headerName = Trim(ws.Cells(1, j).Value)
            
            If headerName <> "" Then
                If Not firstCol Then json = json & ","
                firstCol = False
                
                Dim cellVal As String
                cellVal = ws.Cells(i, j).Value
                
                ' Format dates
                If IsDate(cellVal) Then
                    cellVal = Format(CDate(cellVal), "dd/mm/yyyy")
                End If
                
                json = json & """" & EscapeJSON(headerName) & """: """ & EscapeJSON(CStr(cellVal)) & """"
            End If
        Next j
        
        json = json & "}"
    Next i
    
    json = json & "]"
    
    GetDetailRecords = json
End Function

' Helper: Convert dictionary keys to JSON array
Function DictToJSONArray(dict As Object) As String
    If dict.Count = 0 Then
        DictToJSONArray = "[]"
        Exit Function
    End If
    
    Dim json As String
    json = "["
    
    Dim keys() As Variant, i As Long
    keys = dict.Keys
    
    ' Sort keys alphabetically
    Dim j As Long, temp As Variant
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(i) > keys(j) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    For i = 0 To UBound(keys)
        If i > 0 Then json = json & ","
        json = json & """" & EscapeJSON(CStr(keys(i))) & """"
    Next i
    
    json = json & "]"
    DictToJSONArray = json
End Function

' Helper: Escape special characters for JSON
Function EscapeJSON(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJSON = result
End Function

