Option Explicit

Sub Button1_Click()
    Dim wsPCI As Worksheet, wsShp As Worksheet, wsOutput As Worksheet, wsInputs As Worksheet
    Dim lastRowPCI As Long, i As Long, outputRow As Long
    Dim streetID As String, sectionID As String, key As String
    Dim matchRow As Variant
    Dim latValue As Variant, lonValue As Variant
    Dim shpStreetSecCol As Long, shpLatCol As Long, shpLonCol As Long
    Dim pciStreetIDCol As Long, pciSectionIDCol As Long, diffCol As Long
    
    ' Additional columns for output
    Dim mapillaryURLCol As Long, mapillaryDateCol As Long
    Dim mapillaryLatCol As Long, mapillaryLonCol As Long
    Dim shpLatOutputCol As Long, shpLonOutputCol As Long
    Dim distanceCol As Long
    
    ' Thresholds and filter from cells
    Dim negativeThreshold As Double
    Dim positiveThreshold As Double
    Dim onlyNoWorkHistory As String
    
    ' Get the active worksheet for inputs
    Set wsInputs = ActiveSheet
    
    ' Read values from cells
    negativeThreshold = wsInputs.Range("B4").Value
    positiveThreshold = wsInputs.Range("B7").Value
    onlyNoWorkHistory = UCase(Trim(wsInputs.Range("B10").Value))
    
    ' Validate inputs
    If Not IsNumeric(negativeThreshold) Then
        MsgBox "Invalid negative threshold in cell B4. Please enter a number.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(positiveThreshold) Then
        MsgBox "Invalid positive threshold in cell B7. Please enter a number.", vbExclamation
        Exit Sub
    End If
    
    If onlyNoWorkHistory <> "YES" And onlyNoWorkHistory <> "NO" Then
        MsgBox "Cell B10 must contain 'Yes' or 'No'.", vbExclamation
        Exit Sub
    End If
    
    Dim diffValue As Double
    Dim useConcatForStreetSec As Boolean
    Dim shpStreetIDCol As Long, shpSectionIDCol As Long
    
    ' Columns to copy from PCI Differences
    Dim columnsToInclude As Variant
    columnsToInclude = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14)
    
    Dim outputCol As Long
    Dim sourceCol As Variant
    
    ' Identify worksheets
    On Error Resume Next
    Set wsPCI = ThisWorkbook.Worksheets("PCI Differences")
    Set wsShp = ThisWorkbook.Worksheets("Shapefile Data")
    On Error GoTo 0
    
    If wsPCI Is Nothing Or wsShp Is Nothing Then
        MsgBox "Required sheets not found!", vbCritical
        Exit Sub
    End If
    
    ' Create or clear "Output" sheet
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Worksheets("Output")
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Worksheets.Add
        wsOutput.Name = "Output"
    Else
        wsOutput.Cells.Clear
    End If
    On Error GoTo 0
    
    Application.StatusBar = "Initializing..."
    
    ' Get column numbers in Shapefile Data
    shpStreetSecCol = GetColumnNumber(wsShp, "StreetSec")
    shpLatCol = GetColumnNumber(wsShp, "Lat")
    shpLonCol = GetColumnNumber(wsShp, "Long")
    
    ' If there's no combined "StreetSec" column, we do the 2-column approach
    If shpStreetSecCol = 0 Then
        useConcatForStreetSec = True
        shpStreetIDCol = GetColumnNumber(wsShp, "StreetID")
        shpSectionIDCol = GetColumnNumber(wsShp, "SectionID")
        If shpStreetIDCol = 0 Or shpSectionIDCol = 0 Then
            MsgBox "Required columns (StreetID and SectionID) not found in Shapefile Data.", vbExclamation
            Exit Sub
        End If
    Else
        useConcatForStreetSec = False
    End If
    
    If shpLatCol = 0 Or shpLonCol = 0 Then
        MsgBox "Could not find columns 'Lat' or 'Long' in Shapefile Data.", vbExclamation
        Exit Sub
    End If
    
    ' Get necessary columns in PCI Differences
    pciStreetIDCol = GetColumnNumber(wsPCI, "Street ID")
    pciSectionIDCol = GetColumnNumber(wsPCI, "Section ID")
    diffCol = GetColumnNumber(wsPCI, "Diff")
    
    If pciStreetIDCol = 0 Or pciSectionIDCol = 0 Or diffCol = 0 Then
        MsgBox "Required columns not found in PCI Differences sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Copy headers to Output
    outputCol = 1
    For Each sourceCol In columnsToInclude
        If (CLng(sourceCol) >= 6 And CLng(sourceCol) <= 9) Or _
           (CLng(sourceCol) >= 11 And CLng(sourceCol) <= 12) Then
            wsOutput.Cells(1, outputCol).Value = wsPCI.Cells(1, CLng(sourceCol)).Value & " " & _
                                                wsPCI.Cells(2, CLng(sourceCol)).Value
        Else
            wsOutput.Cells(1, outputCol).Value = wsPCI.Cells(1, CLng(sourceCol)).Value
        End If
        outputCol = outputCol + 1
    Next sourceCol
    
    ' Add columns for shapefile lat/long
    shpLatOutputCol = outputCol
    shpLonOutputCol = shpLatOutputCol + 1
    
    wsOutput.Cells(1, shpLatOutputCol).Value = "Shapefile Lat"
    wsOutput.Cells(1, shpLonOutputCol).Value = "Shapefile Long"
    
    ' Add columns for Mapillary data
    mapillaryURLCol = shpLonOutputCol + 1
    mapillaryDateCol = mapillaryURLCol + 1
    mapillaryLatCol = mapillaryDateCol + 1
    mapillaryLonCol = mapillaryLatCol + 1
    
    wsOutput.Cells(1, mapillaryURLCol).Value = "Mapillary Image URL"
    wsOutput.Cells(1, mapillaryDateCol).Value = "Image Date"
    wsOutput.Cells(1, mapillaryLatCol).Value = "Image Latitude"
    wsOutput.Cells(1, mapillaryLonCol).Value = "Image Longitude"
    
    ' Finally, a column for Distance
    distanceCol = mapillaryLonCol + 1
    wsOutput.Cells(1, distanceCol).Value = "Distance between Image and Shapefile's Coordinates (Miles)"
    
    ' Format header row
    With wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(1, distanceCol))
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
    End With
    
    ' Find last row in PCI Differences
    lastRowPCI = wsPCI.Cells(wsPCI.Rows.Count, pciStreetIDCol).End(xlUp).Row
    outputRow = 2
    
    ' Loop through PCI Differences
    For i = 3 To lastRowPCI
        Application.StatusBar = "Processing row " & i & " of " & lastRowPCI
        If Not IsEmpty(wsPCI.Cells(i, diffCol)) Then
            diffValue = wsPCI.Cells(i, diffCol).Value
            
            ' Condition 1: <= negativeThreshold
            Dim condition1 As Boolean
            condition1 = (diffValue <= negativeThreshold)
            
            ' Condition 2: >= positiveThreshold
            Dim condition2 As Boolean
            If onlyNoWorkHistory = "YES" Then
                ' Must also have M&R columns empty
                condition2 = (diffValue >= positiveThreshold) And _
                             IsEmpty(wsPCI.Cells(i, 8).Value) And _
                             IsEmpty(wsPCI.Cells(i, 9).Value) And _
                             IsEmpty(wsPCI.Cells(i, 10).Value)
            Else
                condition2 = (diffValue >= positiveThreshold)
            End If
            
            ' If row qualifies
            If condition1 Or condition2 Then
                ' Build key (StreetID - SectionID)
                streetID = Trim(CStr(wsPCI.Cells(i, pciStreetIDCol).Value))
                sectionID = Trim(CStr(wsPCI.Cells(i, pciSectionIDCol).Value))
                key = streetID & " - " & sectionID
                
                ' Attempt to find a match in Shapefile Data
                If Not useConcatForStreetSec Then
                    ' If there's a "StreetSec" column
                    matchRow = Application.Match(key, _
                                wsShp.Range(wsShp.Cells(2, shpStreetSecCol), _
                                            wsShp.Cells(wsShp.Rows.Count, shpStreetSecCol)), 0)
                    If Not IsError(matchRow) Then matchRow = matchRow + 1
                Else
                    ' Otherwise, manually search StreetID+SectionID
                    Dim shpRow As Long, lastRowShp As Long, tempKey As String
                    lastRowShp = wsShp.Cells(wsShp.Rows.Count, shpStreetIDCol).End(xlUp).Row
                    matchRow = 0
                    For shpRow = 2 To lastRowShp
                        tempKey = Trim(CStr(wsShp.Cells(shpRow, shpStreetIDCol).Value)) & " - " & _
                                  Trim(CStr(wsShp.Cells(shpRow, shpSectionIDCol).Value))
                        If tempKey = key Then
                            matchRow = shpRow
                            Exit For
                        End If
                    Next shpRow
                    If matchRow = 0 Then matchRow = CVErr(xlErrNA)
                End If
                
                ' If found a matching shapefile row
                If Not IsError(matchRow) Then
                    latValue = wsShp.Cells(matchRow, shpLatCol).Value
                    lonValue = wsShp.Cells(matchRow, shpLonCol).Value
                    
                    If IsNumeric(latValue) And IsNumeric(lonValue) Then
                        ' Copy columns from PCI Differences
                        outputCol = 1
                        For Each sourceCol In columnsToInclude
                            wsOutput.Cells(outputRow, outputCol).Value = _
                                wsPCI.Cells(i, CLng(sourceCol)).Value
                            outputCol = outputCol + 1
                        Next sourceCol
                        
                        ' Write Shapefile Lat/Long
                        wsOutput.Cells(outputRow, shpLatOutputCol).Value = latValue
                        wsOutput.Cells(outputRow, shpLonOutputCol).Value = lonValue
                        
                        ' Mapillary API call
                        Application.StatusBar = "Making API call for row " & i & " of " & lastRowPCI
                        Dim mapillaryResponse As String
                        mapillaryResponse = GetMapillaryResponse(CDbl(latValue), CDbl(lonValue))
                        
                        ' Extract ID, coords, date
                        Dim imageId As String, coordinates As String
                        imageId = ExtractMapillaryID(mapillaryResponse)
                        coordinates = ExtractCoordinates(mapillaryResponse)
                        
                        If imageId <> "" Then
                            ' Write Mapillary link
                            wsOutput.Cells(outputRow, mapillaryURLCol).Value = _
                                "https://www.mapillary.com/app/?focus=photo&pKey=" & imageId
                            
                            ' Make the URL a clickable hyperlink
                            wsOutput.Cells(outputRow, mapillaryURLCol).Hyperlinks.Add _
                                Anchor:=wsOutput.Cells(outputRow, mapillaryURLCol), _
                                Address:=wsOutput.Cells(outputRow, mapillaryURLCol).Value, _
                                TextToDisplay:=wsOutput.Cells(outputRow, mapillaryURLCol).Value
                            
                            ' Write date
                            wsOutput.Cells(outputRow, mapillaryDateCol).Value = _
                                ExtractMapillaryDate(mapillaryResponse)
                            
                            ' Parse mapillary coordinates
                            If coordinates <> "" Then
                                Dim coordArray() As String
                                coordinates = Replace(coordinates, "[", "")
                                coordinates = Replace(coordinates, "]", "")
                                coordArray = Split(Trim(coordinates), ",")
                                
                                ' Mapillary returns [longitude, latitude]
                                If UBound(coordArray) = 1 Then
                                    Dim mapLon As Double, mapLat As Double
                                    
                                    mapLon = CDbl(Trim(coordArray(0)))
                                    mapLat = CDbl(Trim(coordArray(1)))
                                    
                                    ' Place them in the correct columns
                                    wsOutput.Cells(outputRow, mapillaryLonCol).Value = _
                                        Format(mapLon, "0.000000")
                                    wsOutput.Cells(outputRow, mapillaryLatCol).Value = _
                                        Format(mapLat, "0.000000")
                                    
                                    ' Compute the distance
                                    wsOutput.Cells(outputRow, distanceCol).Value = _
                                        HaversineDistanceMiles(CDbl(latValue), CDbl(lonValue), _
                                                               mapLat, mapLon)
                                End If
                            End If
                        Else
                            ' If no image found
                            wsOutput.Cells(outputRow, distanceCol).Value = "N/A"
                        End If
                        
                        outputRow = outputRow + 1
                    End If
                End If
            End If
        End If
        DoEvents
    Next i
    
    ' Optional: highlight negative/positive Diff in Output
    Dim diffColOutput As Long
    diffColOutput = GetColumnNumber(wsOutput, "Diff")
    If diffColOutput > 0 And outputRow > 2 Then
        With wsOutput.Range(wsOutput.Cells(2, diffColOutput), wsOutput.Cells(outputRow - 1, diffColOutput))
            ' Negative in light red
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 200, 200)
            
            ' Positive in light green
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(200, 255, 200)
        End With
    End If
    
    ' If we actually copied rows
    If outputRow > 2 Then
        With wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(outputRow - 1, distanceCol))
            ' Borders
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            
            ' Alignment
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            
            ' Alternate row coloring
            Dim rng As Range
            For Each rng In .Rows
                If rng.Row Mod 2 = 0 Then
                    rng.Interior.Color = RGB(242, 242, 242)
                End If
            Next rng
        End With
        
        ' Autofit
        wsOutput.Columns("A:" & Split(wsOutput.Cells(1, distanceCol).Address, "$")(1)).AutoFit
        wsOutput.Columns(mapillaryURLCol).ColumnWidth = 50  ' widen Mapillary URL
    End If
    
    Application.StatusBar = False
    If outputRow = 2 Then
        MsgBox "No rows found with the specified conditions.", vbInformation
    Else
        MsgBox "Processing complete. Found " & (outputRow - 2) & _
               " rows meeting the threshold(s).", vbInformation
    End If
End Sub

'=======================================================================================
' Function: GetMapillaryResponse
' Purpose: Makes a GET request to the Mapillary API using the given latitude and longitude.
'=======================================================================================
Function GetMapillaryResponse(lat As Double, lon As Double) As String
    Dim httpReq As Object, url As String
    Dim accessToken As String
    
    accessToken = "MLY|9441786265842838|7f6f0c2a2d6a89b3aa725bdd2cb34fd0"
    url = "https://graph.mapillary.com/images?access_token=" & accessToken & _
          "&fields=id,geometry,captured_at&bbox=" & (lon - 0.001) & "," & (lat - 0.001) & "," & _
          (lon + 0.001) & "," & (lat + 0.001) & "&limit=1"
    
    Set httpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error Resume Next
    httpReq.Open "GET", url, False
    httpReq.Send
    
    If Err.Number = 0 Then
        GetMapillaryResponse = httpReq.responseText
    Else
        GetMapillaryResponse = "API Error"
    End If
    On Error GoTo 0
End Function

'=======================================================================================
' Function: ExtractMapillaryID
' Purpose: Extracts the image id from the JSON API response.
'=======================================================================================
Function ExtractMapillaryID(jsonResponse As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, jsonResponse, """id"":""") + 6
    If startPos > 6 Then
        endPos = InStr(startPos, jsonResponse, """")
        If endPos > 0 Then
            ExtractMapillaryID = Mid(jsonResponse, startPos, endPos - startPos)
        End If
    End If
End Function

'=======================================================================================
' Function: ExtractCoordinates
' Purpose: Extracts the coordinates ( [lon, lat] ) from the JSON API response.
'=======================================================================================
Function ExtractCoordinates(jsonResponse As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, jsonResponse, """coordinates"":[") + 14
    If startPos > 14 Then
        endPos = InStr(startPos, jsonResponse, "]")
        If endPos > 0 Then
            ExtractCoordinates = Mid(jsonResponse, startPos, endPos - startPos)
        End If
    End If
End Function

'=======================================================================================
' Function: ExtractMapillaryDate
' Purpose: Extracts and converts the 'captured_at' field from the JSON API response.
'=======================================================================================
Function ExtractMapillaryDate(jsonResponse As String) As String
    Dim startPos As Long, endPos As Long
    Dim dateStr As String, firstChar As String
    Dim ts As Double
    
    startPos = InStr(1, jsonResponse, """captured_at"":")
    If startPos > 0 Then
        startPos = startPos + Len("""captured_at"":")
        Do While Mid(jsonResponse, startPos, 1) = " " Or Mid(jsonResponse, startPos, 1) = vbTab
            startPos = startPos + 1
        Loop
        
        firstChar = Mid(jsonResponse, startPos, 1)
        If firstChar = """" Then
            ' ISO 8601 like "2022-01-01T12:34:56Z"
            startPos = startPos + 1
            endPos = InStr(startPos, jsonResponse, """")
            If endPos > startPos Then
                dateStr = Mid(jsonResponse, startPos, endPos - startPos)
                dateStr = Replace(dateStr, "T", " ")
                dateStr = Replace(dateStr, "Z", "")
                On Error Resume Next
                ExtractMapillaryDate = Format(CDate(dateStr), "yyyy-mm-dd hh:mm:ss")
                If Err.Number <> 0 Then
                    ExtractMapillaryDate = "Invalid Date"
                End If
                On Error GoTo 0
            Else
                ExtractMapillaryDate = "Date Not Found"
            End If
        Else
            ' Could be numeric timestamp in ms
            endPos = startPos
            Do While IsNumeric(Mid(jsonResponse, endPos, 1)) Or Mid(jsonResponse, endPos, 1) = "."
                endPos = endPos + 1
            Loop
            dateStr = Trim(Mid(jsonResponse, startPos, endPos - startPos))
            On Error Resume Next
            ts = CDbl(dateStr)
            ExtractMapillaryDate = Format((ts / 86400000) + 25569, "yyyy-mm-dd hh:mm:ss")
            If Err.Number <> 0 Then
                ExtractMapillaryDate = "Invalid Date"
            End If
            On Error GoTo 0
        End If
    Else
        ExtractMapillaryDate = "Date Not Found"
    End If
End Function

'=======================================================================================
' Function: GetColumnNumber
' Purpose: Returns the column number for the given headerName in row 1 of the specified sheet.
'=======================================================================================
Function GetColumnNumber(ws As Worksheet, headerName As String) As Long
    Dim cell As Range
    For Each cell In ws.Range("1:1")
        If Trim(cell.Value) = headerName Then
            GetColumnNumber = cell.Column
            Exit Function
        End If
    Next cell
    GetColumnNumber = 0
End Function

'=======================================================================================
' Function: HaversineDistanceMiles
' Purpose: Returns the haversine distance (in miles) between two lat/lon points.
'=======================================================================================
Function HaversineDistanceMiles(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    Const RADIUS_EARTH_MILES As Double = 3958.8
    Dim lat1Rad As Double, lat2Rad As Double
    Dim dLat As Double, dLon As Double
    Dim a As Double, c As Double
    
    ' Convert degrees to radians
    lat1Rad = WorksheetFunction.Radians(lat1)
    lat2Rad = WorksheetFunction.Radians(lat2)
    dLat = WorksheetFunction.Radians(lat2 - lat1)
    dLon = WorksheetFunction.Radians(lon2 - lon1)
    
    ' Apply Haversine formula
    a = Sin(dLat / 2) * Sin(dLat / 2) + _
        Cos(lat1Rad) * Cos(lat2Rad) * _
        Sin(dLon / 2) * Sin(dLon / 2)
        
    ' Use regular arctangent function instead
    If a > 1 Then a = 1  ' Prevent domain error
    c = 2 * Atn(Sqr(a) / Sqr(1 - a))
    
    HaversineDistanceMiles = RADIUS_EARTH_MILES * c
End Function

