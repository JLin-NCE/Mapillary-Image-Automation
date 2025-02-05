Sub Button1_Click()
    Dim wsPCI As Worksheet, wsShp As Worksheet, wsOutput As Worksheet
    Dim lastRowPCI As Long, i As Long, outputRow As Long
    Dim streetID As String, sectionID As String, key As String
    Dim matchRow As Variant
    Dim latValue As Variant, lonValue As Variant
    Dim shpStreetSecCol As Long, shpLatCol As Long, shpLonCol As Long
    Dim pciStreetIDCol As Long, pciSectionIDCol As Long, diffCol As Long
    Dim mapillaryURLCol As Long, mapillaryDateCol As Long
    Dim mapillaryLatCol As Long, mapillaryLonCol As Long
    Dim threshold As Variant
    Dim diffValue As Double
    Dim useConcatForStreetSec As Boolean
    Dim shpStreetIDCol As Long, shpSectionIDCol As Long
    Dim columnsToInclude As Variant
    Dim outputCol As Long
    Dim sourceCol As Variant
    Dim col As Long
    
    ' Define columns to include (A to G, K, L, and N)
    columnsToInclude = Array(1, 2, 3, 4, 5, 6, 7, 11, 12, 14)
    
    ' Variables for batching API calls
    Dim apiCallCount As Long
    Dim batchSize As Long
    batchSize = 15   ' Adjust the batch size as needed
    apiCallCount = 0
    
    ' Get user input for threshold
    threshold = Application.InputBox("Enter the minimum difference value to include (NOT absolute):", _
                                    "Difference Threshold", Type:=1)
    If TypeName(threshold) = "Boolean" Then
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
    
    ' Get column numbers from Shapefile Data sheet
    shpStreetSecCol = GetColumnNumber(wsShp, "StreetSec")
    shpLatCol = GetColumnNumber(wsShp, "Lat")
    shpLonCol = GetColumnNumber(wsShp, "Long")
    
    ' If StreetSec column is missing, use concatenated values
    If shpStreetSecCol = 0 Then
        useConcatForStreetSec = True
        shpStreetIDCol = GetColumnNumber(wsShp, "StreetID")
        shpSectionIDCol = GetColumnNumber(wsShp, "SectionID")
        If shpStreetIDCol = 0 Or shpSectionIDCol = 0 Then
            MsgBox "Required columns (StreetID and SectionID) not found in Shapefile Data sheet.", vbExclamation
            Application.StatusBar = False
            Exit Sub
        End If
    Else
        useConcatForStreetSec = False
    End If
    
    ' Validate columns
    If shpLatCol = 0 Or shpLonCol = 0 Then
        MsgBox "Required columns (Lat or Long) not found in Shapefile Data sheet.", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' Get column numbers from PCI Differences sheet
    pciStreetIDCol = GetColumnNumber(wsPCI, "Street ID")
    pciSectionIDCol = GetColumnNumber(wsPCI, "Section ID")
    diffCol = GetColumnNumber(wsPCI, "Diff")
    
    ' Validate required columns
    If pciStreetIDCol = 0 Or pciSectionIDCol = 0 Or diffCol = 0 Then
        MsgBox "Required columns not found in PCI Differences sheet.", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' Copy headers with special handling for certain columns
    outputCol = 1
    For Each sourceCol In columnsToInclude
        
        ' --- START of "combine row1 + row2" logic
        ' We used to do this for columns 6..9 (F..I).
        ' Now let's also do it for columns 11..12 (K..L).
        If (CLng(sourceCol) >= 6 And CLng(sourceCol) <= 9) Or _
           (CLng(sourceCol) >= 11 And CLng(sourceCol) <= 12) Then
            wsOutput.Cells(1, outputCol).value = wsPCI.Cells(1, CLng(sourceCol)).value & " " & _
                                                wsPCI.Cells(2, CLng(sourceCol)).value
        Else
            wsOutput.Cells(1, outputCol).value = wsPCI.Cells(1, CLng(sourceCol)).value
        End If
        ' --- END of "combine row1 + row2" logic
        
        outputCol = outputCol + 1
    Next sourceCol
    
    ' Set up new columns
    mapillaryURLCol = outputCol
    mapillaryDateCol = mapillaryURLCol + 1
    mapillaryLatCol = mapillaryDateCol + 1
    mapillaryLonCol = mapillaryLatCol + 1
    
    ' Add headers for new columns
    wsOutput.Cells(1, mapillaryURLCol).value = "Mapillary Image URL"
    wsOutput.Cells(1, mapillaryDateCol).value = "Image Date"
    wsOutput.Cells(1, mapillaryLatCol).value = "Image Latitude"
    wsOutput.Cells(1, mapillaryLonCol).value = "Image Longitude"
    
    ' Format headers
    With wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(1, mapillaryLonCol))
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30 ' Adjust row height for wrapped text
    End With
    
    ' Find the last row with data
    lastRowPCI = wsPCI.Cells(wsPCI.Rows.Count, pciStreetIDCol).End(xlUp).Row
    outputRow = 2
    
    ' Process rows
    For i = 3 To lastRowPCI
        ' Update status bar
        Application.StatusBar = "Processing row " & i & " of " & lastRowPCI
        
        If Not IsEmpty(wsPCI.Cells(i, diffCol)) Then
            diffValue = wsPCI.Cells(i, diffCol).value
            
            If diffValue >= threshold Then
                ' Get the key from PCI sheet
                streetID = Trim(CStr(wsPCI.Cells(i, pciStreetIDCol).value))
                sectionID = Trim(CStr(wsPCI.Cells(i, pciSectionIDCol).value))
                key = streetID & " - " & sectionID
                
                ' Find matching row in Shapefile Data sheet
                If Not useConcatForStreetSec Then
                    matchRow = Application.Match(key, wsShp.Range(wsShp.Cells(2, shpStreetSecCol), _
                             wsShp.Cells(wsShp.Rows.Count, shpStreetSecCol)), 0)
                    If Not IsError(matchRow) Then
                        matchRow = matchRow + 1
                    End If
                Else
                    ' Build the matching key using concatenated StreetID and SectionID
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
                    If matchRow = 0 Then
                        matchRow = CVErr(xlErrNA)
                    End If
                End If
                
                ' If a match is found, retrieve the latitude and longitude
                If Not IsError(matchRow) Then
                    latValue = wsShp.Cells(matchRow, shpLatCol).value
                    lonValue = wsShp.Cells(matchRow, shpLonCol).value
                    
                    If IsNumeric(latValue) And IsNumeric(lonValue) Then
                        ' Copy specified columns from PCI Differences
                        outputCol = 1
                        For Each sourceCol In columnsToInclude
                            wsOutput.Cells(outputRow, outputCol).value = wsPCI.Cells(i, CLng(sourceCol)).value
                            outputCol = outputCol + 1
                        Next sourceCol
                        
                        ' Make API call
                        Application.StatusBar = "Making API call for row " & i & " of " & lastRowPCI
                        
                        Dim mapillaryResponse As String
                        mapillaryResponse = GetMapillaryResponse(CDbl(latValue), CDbl(lonValue))
                        
                        ' Extract data from response
                        Dim imageId As String, coordinates As String
                        imageId = ExtractMapillaryID(mapillaryResponse)
                        coordinates = ExtractCoordinates(mapillaryResponse)
                        
                        If imageId <> "" Then
                            wsOutput.Cells(outputRow, mapillaryURLCol).value = _
                                "https://www.mapillary.com/app/?focus=photo&pKey=" & imageId
                            wsOutput.Cells(outputRow, mapillaryDateCol).value = ExtractMapillaryDate(mapillaryResponse)
                            
                            ' Handle coordinates - remove any brackets and clean up the format
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
        DoEvents ' Allow UI updates
    Next i
    
    ' Format the output if we have data
    If outputRow > 2 Then
        With wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(outputRow - 1, mapillaryLonCol))
            ' Borders
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            
            ' General formatting
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
        
        ' Autofit columns
        wsOutput.Columns("A:" & Split(wsOutput.Cells(1, mapillaryLonCol).Address, "$")(1)).AutoFit
        
        ' Make URL column wider for better visibility
        wsOutput.Columns(mapillaryURLCol).ColumnWidth = 50
    End If
    
    ' Reset status bar
    Application.StatusBar = False
    
    ' Show completion message
    If outputRow = 2 Then
        MsgBox "No rows found meeting the difference threshold of " & threshold, vbInformation
    Else
        MsgBox "Processing complete. Found " & (outputRow - 2) & " rows meeting the difference threshold.", vbInformation
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
          "&fields=id,geometry,captured_at&bbox=" & lon - 0.001 & "," & lat - 0.001 & "," & _
          lon + 0.001 & "," & lat + 0.001 & "&limit=1"

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
        ' Move past the label
        startPos = startPos + Len("""captured_at"":")
        ' Skip any spaces or tabs
        Do While Mid(jsonResponse, startPos, 1) = " " Or Mid(jsonResponse, startPos, 1) = vbTab
            startPos = startPos + 1
        Loop
        
        firstChar = Mid(jsonResponse, startPos, 1)
        If firstChar = """" Then
            ' Handle ISO 8601 string (e.g., "2021-10-15T13:45:20Z")
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
            ' Handle numeric timestamp (milliseconds since Unix epoch)
            endPos = startPos
            Do While IsNumeric(Mid(jsonResponse, endPos, 1)) Or Mid(jsonResponse, endPos, 1) = "."
                endPos = endPos + 1
            Loop
            dateStr = Trim(Mid(jsonResponse, startPos, endPos - startPos))
            On Error Resume Next
            ts = CDbl(dateStr)
            ' Convert milliseconds to Excel date
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
' Purpose: Returns the column number for the header matching headerName in the first row.
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


