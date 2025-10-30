Option Explicit

' ==========================================
' Portfolio Analytics Data Aggregator
' ==========================================
' This VBA script processes Excel data and creates an optimized JSON file
' with aggregated data, pre-calculated metrics, and proper date formatting
' 
' Instructions:
' 1. Open your Excel file with the raw data
' 2. Press Alt+F11 to open VBA Editor
' 3. Insert > Module, then paste this code
' 4. Update the CONFIG section below with your settings
' 5. Run the "ExportPortfolioDataToJSON" macro
' ==========================================

' ==========================================
' CONFIGURATION - Update these values
' ==========================================
Const DATA_SHEET_NAME As String = "Sheet1"        ' Name of your data sheet
Const START_ROW As Integer = 2                     ' First row of data (after headers)
Const OUTPUT_FOLDER As String = ""                 ' Leave empty to save in same folder as workbook
Const OUTPUT_FILENAME As String = "portfolioData.json"
Const MAX_DETAIL_RECORDS As Long = 5000           ' Limit detail records for performance
' ==========================================

' Main export function
Sub ExportPortfolioDataToJSON()
    On Error GoTo ErrorHandler
    
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Debug.Print "=========================================="
    Debug.Print "Portfolio Data Aggregator Started"
    Debug.Print "=========================================="
    
    ' Get the data sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    
    ' Find last row and last column
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Debug.Print "Found " & (lastRow - START_ROW + 1) & " records"
    
    ' Read headers
    Dim headers() As String
    ReDim headers(1 To lastCol)
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To lastCol
        headers(i) = Trim(ws.Cells(1, i).Value)
        colMap(headers(i)) = i
    Next i
    
    ' Validate required columns
    If Not ValidateColumns(colMap) Then
        MsgBox "Missing required columns. Please check the data.", vbCritical
        Exit Sub
    End If
    
    ' Read and process data
    Dim rawData() As Variant
    rawData = ws.Range(ws.Cells(START_ROW, 1), ws.Cells(lastRow, lastCol)).Value
    
    Debug.Print "Processing data..."
    
    ' Create aggregated data structures
    Dim metrics As Object
    Set metrics = CalculateMetrics(rawData, colMap)
    
    Dim chartData As Object
    Set chartData = AggregateChartData(rawData, colMap)
    
    Dim filterOptions As Object
    Set filterOptions = ExtractFilterOptions(rawData, colMap)
    
    ' Limit detail records for performance
    Dim detailRecords As Long
    detailRecords = Application.Min(lastRow - START_ROW + 1, MAX_DETAIL_RECORDS)
    
    Debug.Print "Creating JSON with " & detailRecords & " detail records..."
    
    ' Build JSON
    Dim json As String
    json = BuildJSON(rawData, colMap, metrics, chartData, filterOptions, detailRecords, headers)
    
    ' Save to file
    Dim outputPath As String
    If OUTPUT_FOLDER = "" Then
        outputPath = ThisWorkbook.Path & "\" & OUTPUT_FILENAME
    Else
        outputPath = OUTPUT_FOLDER & "\" & OUTPUT_FILENAME
    End If
    
    SaveTextToFile json, outputPath
    
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Debug.Print "=========================================="
    Debug.Print "Export completed in " & Format(elapsedTime, "0.00") & " seconds"
    Debug.Print "File saved to: " & outputPath
    Debug.Print "=========================================="
    
    MsgBox "Data exported successfully!" & vbCrLf & vbCrLf & _
           "File: " & outputPath & vbCrLf & _
           "Records: " & detailRecords & vbCrLf & _
           "Time: " & Format(elapsedTime, "0.00") & "s", vbInformation, "Export Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical
    Debug.Print "Error: " & Err.Description
End Sub

' Validate required columns exist
Function ValidateColumns(colMap As Object) As Boolean
    Dim requiredCols As Variant
    requiredCols = Array("Region", "Facility_ID", "Relationship_Name", _
                        "Product_Program", "Fac_Amount", "Direct_OS", _
                        "Maturity_Date", "CA_Expiration_Date")
    
    Dim col As Variant
    For Each col In requiredCols
        If Not colMap.Exists(col) Then
            Debug.Print "Missing required column: " & col
            ValidateColumns = False
            Exit Function
        End If
    Next col
    
    ValidateColumns = True
End Function

' Calculate all metrics
Function CalculateMetrics(data As Variant, colMap As Object) As Object
    Dim metrics As Object
    Set metrics = CreateObject("Scripting.Dictionary")
    
    Dim totalFacAmount As Double
    Dim totalOSUC As Double
    Dim relationships As Object
    Dim facilities As Object
    Dim cas As Object
    
    Set relationships = CreateObject("Scripting.Dictionary")
    Set facilities = CreateObject("Scripting.Dictionary")
    Set cas = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        ' Sum amounts
        totalFacAmount = totalFacAmount + GetNumericValue(data(i, colMap("Fac_Amount")))
        totalOSUC = totalOSUC + GetNumericValue(data(i, colMap("Total_Comm_Exposure")))
        
        ' Count uniques
        Dim relName As String
        relName = GetStringValue(data(i, colMap("Relationship_Name")))
        If relName <> "" And Not relationships.Exists(relName) Then
            relationships(relName) = 1
        End If
        
        Dim facID As String
        facID = GetStringValue(data(i, colMap("Facility_ID")))
        If facID <> "" And Not facilities.Exists(facID) Then
            facilities(facID) = 1
        End If
        
        Dim caID As String
        caID = GetStringValue(data(i, colMap("CAID")))
        If caID <> "" And Not cas.Exists(caID) Then
            cas(caID) = 1
        End If
    Next i
    
    metrics("totalRelationships") = relationships.Count
    metrics("totalFacilities") = facilities.Count
    metrics("totalCAs") = cas.Count
    metrics("totalFacAmount") = totalFacAmount
    metrics("totalOSUC") = totalOSUC
    
    Debug.Print "Metrics: " & relationships.Count & " relationships, " & _
                facilities.Count & " facilities, " & cas.Count & " CAs"
    
    Set CalculateMetrics = metrics
End Function

' Aggregate data for charts
Function AggregateChartData(data As Variant, colMap As Object) As Object
    Dim chartData As Object
    Set chartData = CreateObject("Scripting.Dictionary")
    
    ' Expiration timeline by month
    Dim expirationByMonth As Object
    Set expirationByMonth = CreateObject("Scripting.Dictionary")
    
    ' CA expiration by month
    Dim caExpirationByMonth As Object
    Set caExpirationByMonth = CreateObject("Scripting.Dictionary")
    
    ' Regional distribution
    Dim regionCounts As Object
    Set regionCounts = CreateObject("Scripting.Dictionary")
    
    ' Top relationships by amount
    Dim relationshipAmounts As Object
    Set relationshipAmounts = CreateObject("Scripting.Dictionary")
    
    ' Top relationships by Direct OS
    Dim relationshipDirectOS As Object
    Set relationshipDirectOS = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim dateVal As Date
    Dim monthKey As String
    Dim region As String
    Dim relName As String
    Dim amount As Double
    
    For i = 1 To UBound(data, 1)
        ' Maturity Date aggregation
        dateVal = GetDateValue(data(i, colMap("Maturity_Date")))
        If dateVal > 0 Then
            monthKey = Format(dateVal, "yyyy-mm")
            If expirationByMonth.Exists(monthKey) Then
                expirationByMonth(monthKey) = expirationByMonth(monthKey) + 1
            Else
                expirationByMonth(monthKey) = 1
            End If
        End If
        
        ' CA Expiration aggregation
        dateVal = GetDateValue(data(i, colMap("CA_Expiration_Date")))
        If dateVal > 0 Then
            monthKey = Format(dateVal, "yyyy-mm")
            If caExpirationByMonth.Exists(monthKey) Then
                caExpirationByMonth(monthKey) = caExpirationByMonth(monthKey) + 1
            Else
                caExpirationByMonth(monthKey) = 1
            End If
        End If
        
        ' Regional distribution (count facilities expiring in next 90 days)
        dateVal = GetDateValue(data(i, colMap("Maturity_Date")))
        If dateVal > 0 And dateVal <= Date + 90 And dateVal >= Date Then
            region = GetStringValue(data(i, colMap("Region")))
            If region <> "" Then
                If regionCounts.Exists(region) Then
                    regionCounts(region) = regionCounts(region) + 1
                Else
                    regionCounts(region) = 1
                End If
            End If
        End If
        
        ' Relationship amounts
        relName = GetStringValue(data(i, colMap("Relationship_Name")))
        If relName <> "" Then
            amount = GetNumericValue(data(i, colMap("Fac_Amount")))
            If relationshipAmounts.Exists(relName) Then
                relationshipAmounts(relName) = relationshipAmounts(relName) + amount
            Else
                relationshipAmounts(relName) = amount
            End If
            
            amount = GetNumericValue(data(i, colMap("Direct_OS")))
            If relationshipDirectOS.Exists(relName) Then
                relationshipDirectOS(relName) = relationshipDirectOS(relName) + amount
            Else
                relationshipDirectOS(relName) = amount
            End If
        End If
    Next i
    
    ' Convert to arrays and sort
    Set chartData("expirationByMonth") = SortDictionaryByKey(expirationByMonth)
    Set chartData("caExpirationByMonth") = SortDictionaryByKey(caExpirationByMonth)
    Set chartData("regionCounts") = regionCounts
    Set chartData("topRelationshipsByAmount") = GetTopN(relationshipAmounts, 20)
    Set chartData("topRelationshipsByDirectOS") = GetTopN(relationshipDirectOS, 20)
    
    Debug.Print "Chart data aggregated: " & expirationByMonth.Count & " months, " & _
                regionCounts.Count & " regions, " & relationshipAmounts.Count & " relationships"
    
    Set AggregateChartData = chartData
End Function

' Extract unique filter options
Function ExtractFilterOptions(data As Variant, colMap As Object) As Object
    Dim filterOptions As Object
    Set filterOptions = CreateObject("Scripting.Dictionary")
    
    Dim regions As Object
    Dim products As Object
    Dim facilityTypes As Object
    Dim commitments As Object
    Dim managementStatus As Object
    Dim teamLeads As Object
    Dim underwriters As Object
    
    Set regions = CreateObject("Scripting.Dictionary")
    Set products = CreateObject("Scripting.Dictionary")
    Set facilityTypes = CreateObject("Scripting.Dictionary")
    Set commitments = CreateObject("Scripting.Dictionary")
    Set managementStatus = CreateObject("Scripting.Dictionary")
    Set teamLeads = CreateObject("Scripting.Dictionary")
    Set underwriters = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim val As String
    
    For i = 1 To UBound(data, 1)
        ' Regions
        val = GetStringValue(data(i, colMap("Region")))
        If val <> "" And Not regions.Exists(val) Then regions(val) = 1
        
        ' Products
        val = GetStringValue(data(i, colMap("Product_Program")))
        If val <> "" And Not products.Exists(val) Then products(val) = 1
        
        ' Facility Types
        If colMap.Exists("Facility_Type") Then
            val = GetStringValue(data(i, colMap("Facility_Type")))
            If val <> "" And Not facilityTypes.Exists(val) Then facilityTypes(val) = 1
        End If
        
        ' Commitments
        If colMap.Exists("Committed/Uncommitted") Then
            val = GetStringValue(data(i, colMap("Committed/Uncommitted")))
            If val <> "" And Not commitments.Exists(val) Then commitments(val) = 1
        End If
        
        ' Management Status
        If colMap.Exists("Management_Status") Then
            val = GetStringValue(data(i, colMap("Management_Status")))
            If val <> "" And Not managementStatus.Exists(val) Then managementStatus(val) = 1
        End If
        
        ' Team Leads
        If colMap.Exists("Underwriting_Team_Lead") Then
            val = GetStringValue(data(i, colMap("Underwriting_Team_Lead")))
            If val <> "" And Not teamLeads.Exists(val) Then teamLeads(val) = 1
        End If
        
        ' Underwriters
        If colMap.Exists("Lead_Underwriter") Then
            val = GetStringValue(data(i, colMap("Lead_Underwriter")))
            If val <> "" And Not underwriters.Exists(val) Then underwriters(val) = 1
        End If
    Next i
    
    Set filterOptions("regions") = DictKeysToSortedArray(regions)
    Set filterOptions("products") = DictKeysToSortedArray(products)
    Set filterOptions("facilityTypes") = DictKeysToSortedArray(facilityTypes)
    Set filterOptions("commitments") = DictKeysToSortedArray(commitments)
    Set filterOptions("managementStatus") = DictKeysToSortedArray(managementStatus)
    Set filterOptions("teamLeads") = DictKeysToSortedArray(teamLeads)
    Set filterOptions("underwriters") = DictKeysToSortedArray(underwriters)
    
    Debug.Print "Filter options extracted"
    
    Set ExtractFilterOptions = filterOptions
End Function

' Build complete JSON structure
Function BuildJSON(data As Variant, colMap As Object, metrics As Object, _
                   chartData As Object, filterOptions As Object, _
                   maxRecords As Long, headers As Variant) As String
    
    Dim json As String
    json = "{" & vbCrLf
    
    ' Metadata
    json = json & "  ""metadata"": {" & vbCrLf
    json = json & "    ""generatedAt"": """ & Format(Now, "yyyy-mm-dd hh:nn:ss") & """," & vbCrLf
    json = json & "    ""totalRecords"": " & UBound(data, 1) & "," & vbCrLf
    json = json & "    ""detailRecords"": " & maxRecords & "," & vbCrLf
    json = json & "    ""version"": ""1.0""" & vbCrLf
    json = json & "  }," & vbCrLf
    
    ' Metrics
    json = json & "  ""metrics"": {" & vbCrLf
    json = json & "    ""totalRelationships"": " & metrics("totalRelationships") & "," & vbCrLf
    json = json & "    ""totalFacilities"": " & metrics("totalFacilities") & "," & vbCrLf
    json = json & "    ""totalCAs"": " & metrics("totalCAs") & "," & vbCrLf
    json = json & "    ""totalFacAmount"": " & FormatNumber(metrics("totalFacAmount"), 2, vbFalse, vbFalse, vbFalse) & "," & vbCrLf
    json = json & "    ""totalOSUC"": " & FormatNumber(metrics("totalOSUC"), 2, vbFalse, vbFalse, vbFalse) & vbCrLf
    json = json & "  }," & vbCrLf
    
    ' Chart data
    json = json & "  ""chartData"": {" & vbCrLf
    json = json & "    ""expirationByMonth"": " & DictToJSONObject(chartData("expirationByMonth")) & "," & vbCrLf
    json = json & "    ""caExpirationByMonth"": " & DictToJSONObject(chartData("caExpirationByMonth")) & "," & vbCrLf
    json = json & "    ""regionCounts"": " & DictToJSONObject(chartData("regionCounts")) & "," & vbCrLf
    json = json & "    ""topRelationshipsByAmount"": " & DictToJSONObject(chartData("topRelationshipsByAmount")) & "," & vbCrLf
    json = json & "    ""topRelationshipsByDirectOS"": " & DictToJSONObject(chartData("topRelationshipsByDirectOS")) & vbCrLf
    json = json & "  }," & vbCrLf
    
    ' Filter options
    json = json & "  ""filterOptions"": {" & vbCrLf
    json = json & "    ""regions"": " & ArrayToJSONArray(filterOptions("regions")) & "," & vbCrLf
    json = json & "    ""products"": " & ArrayToJSONArray(filterOptions("products")) & "," & vbCrLf
    json = json & "    ""facilityTypes"": " & ArrayToJSONArray(filterOptions("facilityTypes")) & "," & vbCrLf
    json = json & "    ""commitments"": " & ArrayToJSONArray(filterOptions("commitments")) & "," & vbCrLf
    json = json & "    ""managementStatus"": " & ArrayToJSONArray(filterOptions("managementStatus")) & "," & vbCrLf
    json = json & "    ""teamLeads"": " & ArrayToJSONArray(filterOptions("teamLeads")) & "," & vbCrLf
    json = json & "    ""underwriters"": " & ArrayToJSONArray(filterOptions("underwriters")) & vbCrLf
    json = json & "  }," & vbCrLf
    
    ' Detail records (limited for performance)
    json = json & "  ""detailData"": [" & vbCrLf
    
    Dim recordCount As Long
    recordCount = Application.Min(UBound(data, 1), maxRecords)
    
    For i = 1 To recordCount
        json = json & "    {"
        
        ' Add each column
        Dim j As Long
        For j = 1 To UBound(headers)
            Dim colName As String
            colName = headers(j)
            
            json = json & """" & colName & """: "
            
            ' Format value based on column type
            If InStr(colName, "Date") > 0 Then
                json = json & """" & FormatDateToISO(GetDateValue(data(i, j))) & """"
            ElseIf IsNumeric(data(i, j)) And colName <> "Facility_ID" And colName <> "CAID" Then
                json = json & FormatNumber(GetNumericValue(data(i, j)), 2, vbFalse, vbFalse, vbFalse)
            Else
                json = json & """" & EscapeJSON(GetStringValue(data(i, j))) & """"
            End If
            
            If j < UBound(headers) Then json = json & ", "
        Next j
        
        json = json & "}"
        If i < recordCount Then json = json & ","
        json = json & vbCrLf
        
        ' Progress indicator
        If i Mod 500 = 0 Then
            Debug.Print "Processed " & i & " of " & recordCount & " records..."
        End If
    Next i
    
    json = json & "  ]" & vbCrLf
    json = json & "}"
    
    BuildJSON = json
End Function

' ==========================================
' HELPER FUNCTIONS
' ==========================================

Function GetStringValue(val As Variant) As String
    If IsEmpty(val) Or IsNull(val) Then
        GetStringValue = ""
    Else
        GetStringValue = Trim(CStr(val))
    End If
End Function

Function GetNumericValue(val As Variant) As Double
    On Error Resume Next
    GetNumericValue = CDbl(val)
    If Err.Number <> 0 Then GetNumericValue = 0
    On Error GoTo 0
End Function

Function GetDateValue(val As Variant) As Date
    On Error Resume Next
    If IsDate(val) Then
        GetDateValue = CDate(val)
    Else
        GetDateValue = 0
    End If
    On Error GoTo 0
End Function

Function FormatDateToISO(dateVal As Date) As String
    If dateVal = 0 Then
        FormatDateToISO = ""
    Else
        FormatDateToISO = Format(dateVal, "yyyy-mm-dd")
    End If
End Function

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

Function DictKeysToSortedArray(dict As Object) As Variant
    If dict.Count = 0 Then
        DictKeysToSortedArray = Array()
        Exit Function
    End If
    
    Dim arr() As String
    ReDim arr(0 To dict.Count - 1)
    
    Dim i As Long
    Dim key As Variant
    i = 0
    For Each key In dict.Keys
        arr(i) = CStr(key)
        i = i + 1
    Next key
    
    ' Simple bubble sort
    Dim j As Long
    Dim temp As String
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    DictKeysToSortedArray = arr
End Function

Function ArrayToJSONArray(arr As Variant) As String
    If Not IsArray(arr) Then
        ArrayToJSONArray = "[]"
        Exit Function
    End If
    
    If UBound(arr) < LBound(arr) Then
        ArrayToJSONArray = "[]"
        Exit Function
    End If
    
    Dim result As String
    result = "["
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        result = result & """" & EscapeJSON(CStr(arr(i))) & """"
        If i < UBound(arr) Then result = result & ", "
    Next i
    
    result = result & "]"
    ArrayToJSONArray = result
End Function

Function DictToJSONObject(dict As Object) As String
    If dict.Count = 0 Then
        DictToJSONObject = "{}"
        Exit Function
    End If
    
    Dim result As String
    result = "{"
    
    Dim key As Variant
    Dim first As Boolean
    first = True
    
    For Each key In dict.Keys
        If Not first Then result = result & ", "
        result = result & """" & EscapeJSON(CStr(key)) & """: "
        
        If IsNumeric(dict(key)) Then
            result = result & FormatNumber(dict(key), 2, vbFalse, vbFalse, vbFalse)
        Else
            result = result & """" & EscapeJSON(CStr(dict(key))) & """"
        End If
        
        first = False
    Next key
    
    result = result & "}"
    DictToJSONObject = result
End Function

Function SortDictionaryByKey(dict As Object) As Object
    Dim sorted As Object
    Set sorted = CreateObject("Scripting.Dictionary")
    
    If dict.Count = 0 Then
        Set SortDictionaryByKey = sorted
        Exit Function
    End If
    
    ' Get keys and sort
    Dim keys() As String
    ReDim keys(0 To dict.Count - 1)
    
    Dim i As Long
    Dim key As Variant
    i = 0
    For Each key In dict.Keys
        keys(i) = CStr(key)
        i = i + 1
    Next key
    
    ' Bubble sort
    Dim j As Long
    Dim temp As String
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(i) > keys(j) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    ' Add to sorted dictionary
    For i = 0 To UBound(keys)
        sorted(keys(i)) = dict(keys(i))
    Next i
    
    Set SortDictionaryByKey = sorted
End Function

Function GetTopN(dict As Object, n As Long) As Object
    Dim topN As Object
    Set topN = CreateObject("Scripting.Dictionary")
    
    If dict.Count = 0 Then
        Set GetTopN = topN
        Exit Function
    End If
    
    ' Convert to arrays
    Dim keys() As String
    Dim values() As Double
    ReDim keys(0 To dict.Count - 1)
    ReDim values(0 To dict.Count - 1)
    
    Dim i As Long
    Dim key As Variant
    i = 0
    For Each key In dict.Keys
        keys(i) = CStr(key)
        values(i) = dict(key)
        i = i + 1
    Next key
    
    ' Sort by value (descending)
    Dim j As Long
    Dim tempKey As String
    Dim tempVal As Double
    
    For i = 0 To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(i) < values(j) Then
                tempVal = values(i)
                values(i) = values(j)
                values(j) = tempVal
                
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i
    
    ' Take top N
    Dim limit As Long
    limit = Application.Min(n, dict.Count)
    
    For i = 0 To limit - 1
        topN(keys(i)) = values(i)
    Next i
    
    Set GetTopN = topN
End Function

Sub SaveTextToFile(text As String, filePath As String)
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(filePath, True, False)
    
    file.Write text
    file.Close
    
    Set file = Nothing
    Set fso = Nothing
End Sub

