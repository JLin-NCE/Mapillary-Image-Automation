Option Explicit

Sub Button1_Click()
    Dim wsPCI As Worksheet, wsShp As Worksheet, wsOutput As Worksheet
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
    
    ' Thresholds
    Dim negativeThreshold As Variant    ' Include if Diff <= negativeThreshold
    Dim positiveThreshold As Variant    ' Include if Diff >= positiveThreshold
    
    ' Yes/No input to require no-work-history for rows above the positive threshold
    Dim onlyNoWorkHistory As String
    
    Dim diffValue As Double
    Dim useConcatForStreetSec As Boolean
    Dim shpStreetIDCol As Long, shpSectionIDCol As Long
    
    ' We now include columns 8,9,10 (M&R Date, M&R Treatment Name, M&R PCI)
    ' along with A..G (1..7), K,L,N (11,12,14)
    Dim columnsToInclude As Variant
    columnsToInclude = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14)
    
    Dim outputCol As Long
    Dim sourceCol As Variant
    
    ' 1) Ask user for negative threshold
    negativeThreshold = Application.InputBox( _
        Prompt:="Enter the negative threshold (rows = this value will be included):", _
        Title:="Negative Threshold", Type:=1)
    If TypeName(negativeThreshold) = "Boolean" Then Exit Sub  ' Cancel
    
    ' 2) Ask user for positive threshold
    positiveThreshold = Application.InputBox( _
        Prompt:="Enter the positive threshold (rows = this value will be included):", _
        Title:="Positive Threshold", Type:=1)
    If TypeName(positiveThreshold) = "Boolean" Then Exit Sub  ' Cancel
    
    ' 3) Ask "Yes/No" if we only want no-work-history rows for the positive threshold
    onlyNoWorkHistory = Application.InputBox( _
        Prompt:="Do you want to ONLY include rows above the positive threshold that have NO work history? (Yes/No)", _
        Title:="Filter By Work History?", Type:=2)
    If TypeName(onlyNoWorkHistory) = "Boolean" Then Exit Sub  ' Cancel
    
    onlyNoWorkHistory = UCase(Trim(onlyNoWorkHistory))
    If onlyNoWorkHistory <> "YES" And onlyNoWorkHistory <> "NO" Then
        MsgBox "Please type 'Yes' or 'No'.", vbExclamation
        Exit Sub
    End If
    
    ' Set worksheets
    On Error Resume Next
    Set wsPCI = ThisWorkbook.Worksheets("PCI Differences")
    Set wsShp = ThisWorkbook.Worksheets("Shapefile Data")
    On Error GoTo 0
    
    If wsPCI Is Nothing Or wsShp Is Nothing Then
        MsgBox "Required sheets not found!", vbCritical
        Exit Sub
    End If
    
    ' Create or clear Output sheet
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Worksheets("Output")
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Worksheets.Add
        wsOutput.name = "Output"
    Else
        wsOutput.Cells.Clear
    End If
    On Error GoTo 0
    
    Application.StatusBar = "Initializing..."
    
    ' Get column numbers from Shapefile Data
    shpStreetSecCol = GetColumnNumber(wsShp, "StreetSec")
    shpLatCol = GetColumnNumber(wsShp, "Lat")
    shpLonCol = GetColumnNumber(wsShp, "Long")
    
    ' If StreetSec not found, we must concatenate StreetID+SectionID
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
    
    ' Make sure we have Lat/Long
    If shpLatCol = 0 Or shpLonCol = 0 Then
        MsgBox "Could not find columns 'Lat' or 'Long' in Shapefile Data.", vbExclamation
        Exit Sub
    End If
    
    ' Get PCI Differences columns: Street ID, Section ID, Diff
    pciStreetIDCol = GetColumnNumber(wsPCI, "Street ID")
    pciSectionIDCol = GetColumnNumber(wsPCI, "Section ID")
    diffCol = GetColumnNumber(wsPCI, "Diff")
    
    If pciStreetIDCol = 0 Or pciSectionIDCol = 0 Or diffCol = 0 Then
        MsgBox "Required columns not found in PCI Differences sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Copy headers
    outputCol = 1
    For Each sourceCol In columnsToInclude
        ' Combine row 1 + row 2 for columns 6..9 and 11..12 if they have multi-row headers
        If (CLng(sourceCol) >= 6 And CLng(sourceCol) <= 9) Or _
           (CLng(sourceCol) >= 11 And CLng(sourceCol) <= 12) Then
            wsOutput.Cells(1, outputCol).value = wsPCI.Cells(1, CLng(sourceCol)).value & " " & _
                                                wsPCI.Cells(2, CLng(sourceCol)).value
        Else
            wsOutput.Cells(1, outputCol).value = wsPCI.Cells(1, CLng(sourceCol)).value
        End If
        outputCol = outputCol + 1
    Next sourceCol
    
    ' Add two columns for Shapefile Lat/Long
    shpLatOutputCol = outputCol
    shpLonOutputCol = shpLatOutputCol + 1
    
    wsOutput.Cells(1, shpLatOutputCol).value = "Shapefile Lat"
    wsOutput.Cells(1, shpLonOutputCol).value = "Shapefile Long"
    
    ' Then add columns for Mapillary data
    mapillaryURLCol = shpLonOutputCol + 1
    mapillaryDateCol = mapillaryURLCol + 1
    mapillaryLatCol = mapillaryDateCol + 1
    mapillaryLonCol = mapillaryLatCol + 1
    
    wsOutput.Cells(1, mapillaryURLCol).value = "Mapillary Image URL"
    wsOutput.Cells(1, mapillaryDateCol).value = "Image Date"
    wsOutput.Cells(1, mapillaryLatCol).value = "Image Latitude"
    wsOutput.Cells(1, mapillaryLonCol).value = "Image Longitude"
    
    ' Format header row
    With wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(1, mapillaryLonCol))
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
            diffValue = wsPCI.Cells(i, diffCol).value
            
            ' Condition 1: Diff <= negativeThreshold
            Dim condition1 As Boolean
            condition1 = (diffValue <= negativeThreshold)
            
            ' Condition 2: Diff >= positiveThreshold
            Dim condition2 As Boolean
            If onlyNoWorkHistory = "YES" Then
                ' Must also have H, I, J empty
                condition2 = (diffValue >= positiveThreshold) And _
                             IsEmpty(wsPCI.Cells(i, 8).value) And _
                             IsEmpty(wsPCI.Cells(i, 9).value) And _
                             IsEmpty(wsPCI.Cells(i, 10).value)
            Else
                condition2 = (diffValue >= positiveThreshold)
            End If
            
            If condition1 Or condition2 Then
                ' Build key StreetID - SectionID
                streetID = Trim(CStr(wsPCI.Cells(i, pciStreetIDCol).value))
                sectionID = Trim(CStr(wsPCI.Cells(i, pciSectionIDCol).value))
                key = streetID & " - " & sectionID
                
                ' Match in Shapefile Data
                If Not useConcatForStreetSec Then
                    matchRow = Application.Match(key, _
                                wsShp.Range(wsShp.Cells(2, shpStreetSecCol), _
                                            wsShp.Cells(wsShp.Rows.Count, shpStreetSecCol)), 0)
                    If Not IsError(matchRow) Then matchRow = matchRow + 1
                Else
                    Dim shpRow As Long, lastRowShp As Long, tempKey As String
                    lastRowShp = wsShp.Cells(wsShp.Rows.Count, shpStreetIDCol).End(xlUp).Row
                    matchRow = 0
                    For shpRow = 2 To lastRowShp
                        tempKey = Trim(CStr(wsShp.Cells(shpRow, shpStreetIDCol).value)) & " - " & _
                                  Trim(CStr(wsShp.Cells(shpRow, shpSectionIDCol).value))
                        If tempKey = key Then
                            matchRow = shpRow
                            Exit For
                        End If
                    Next shpRow
                    If matchRow = 0 Then matchRow = CVErr(xlErrNA)
                End If
                
                ' If we found a matching shapefile row
                If Not IsError(matchRow) Then
                    latValue = wsShp.Cells(matchRow, shpLatCol).value
                    lonValue = wsShp.Cells(matchRow, shpLonCol).value
                    
                    If IsNumeric(latValue) And IsNumeric(lonValue) Then
                        ' Copy the requested columns from PCI Differences
                        Dim colIndex As Long
                        outputCol = 1
                        For Each sourceCol In columnsToInclude
                            wsOutput.Cells(outputRow, outputCol).value = _
                                wsPCI.Cells(i, CLng(sourceCol)).value
                            outputCol = outputCol + 1
                        Next sourceCol
                        
                        ' Put shapefile Lat/Long
                        wsOutput.Cells(outputRow, shpLatOutputCol).value = latValue
                        wsOutput.Cells(outputRow, shpLonOutputCol).value = lonValue
                        
                        ' Mapillary API call
                        Application.StatusBar = "Making API call for row " & i & " of " & lastRowPCI
                        Dim mapillaryResponse As String
                        mapillaryResponse = GetMapillaryResponse(CDbl(latValue), CDbl(lonValue))
                        
                        ' Extract ID, coords, date
                        Dim imageId As String, coordinates As String
                        imageId = ExtractMapillaryID(mapillaryResponse)
                        coordinates = ExtractCoordinates(mapillaryResponse)
                        
                        If imageId <> "" Then
                            wsOutput.Cells(outputRow, mapillaryURLCol).value = _
                                "https://www.mapillary.com/app/?focus=photo&pKey=" & imageId
                            
                            ' Hyperlink
                            wsOutput.Cells(outputRow, mapillaryURLCol).Hyperlinks.Add _
                                Anchor:=wsOutput.Cells(outputRow, mapillaryURLCol), _
                                Address:=wsOutput.Cells(outputRow, mapillaryURLCol).value, _
                                TextToDisplay:=wsOutput.Cells(outputRow, mapillaryURLCol).value
                            
                            ' Date
                            wsOutput.Cells(outputRow, mapillaryDateCol).value = _
                                ExtractMapillaryDate(mapillaryResponse)
                            
                            ' Mapillary lon/lat
                            If coordinates <> "" Then
                                Dim coordArray() As String
                                coordinates = Replace(coordinates, "[", "")
                                coordinates = Replace(coordinates, "]", "")
                                coordArray = Split(Trim(coordinates), ",")
                                If UBound(coordArray) = 1 Then
                                    wsOutput.Cells(outputRow, mapillaryLonCol).value = _
                                        Format(CDbl(Trim(coordArray(0))), "0.000000")
                                    wsOutput.Cells(outputRow, mapillaryLatCol).value = _
                                        Format(CDbl(Trim(coordArray(1))), "0.000000")
                                End If
                            End If
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
        With wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(outputRow - 1, mapillaryLonCol))
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
        wsOutput.Columns("A:" & Split(wsOutput.Cells(1, mapillaryLonCol).Address, "$")(1)).AutoFit
        ' Widen the Mapillary URL column
        wsOutput.Columns(mapillaryURLCol).ColumnWidth = 50
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
' Purpose: Extracts the coordinates (longitude, latitude) from the JSON API response.
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
            ' Numeric timestamp in ms
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
        If Trim(cell.value) = headerName Then
            GetColumnNumber = cell.Column
            Exit Function
        End If
    Next cell
    GetColumnNumber = 0
End Function


