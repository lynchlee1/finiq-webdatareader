Option Explicit

' FnGuide CompanyGuide: financial statements + financial ratios
' Paste this module into Excel VBA, then run DownloadFnguideTables.

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const BASE_URL As String = "https://comp.fnguide.com/SVO2/ASP/"
Private Const REQUEST_DELAY_MILLISECONDS As Long = 300
Private Const REQUEST_RETRY_COUNT As Long = 6
Private Const REQUEST_RETRY_DELAY_MILLISECONDS As Long = 1000

Private gStep As String
Private gHasRequested As Boolean
Private gCompanyName As String

Private gRegExpDefault As Object
Private gRegExpRow As Object
Private gRegExpCell As Object
Private gRegExpSpan As Object

Public Sub DownloadFnguideTables()
    Dim wsMain As Worksheet
    Dim wsFnGuide As Worksheet
    Dim code As String
    Dim reportGbVal As String
    Dim reportGb As String
    Dim reportName As String
    Dim originalCalculation As Long
    Dim lastRow As Long
    Dim r As Long
    Dim startCol As Long
    Dim totalCols As Long
    Dim companyCount As Long

    ' 1. Check/get MAIN sheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = ThisWorkbook

    On Error Resume Next
    Set wsMain = wb.Worksheets("MAIN")
    On Error GoTo 0
    If wsMain Is Nothing Then
        ShowMsgBox "MAIN sheet was not found. Create MAIN sheet and enter stock codes starting at B3.", vbCritical
        Exit Sub
    End If

    lastRow = wsMain.Cells(wsMain.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        ShowMsgBox "No stock codes found in column B of MAIN starting at row 3.", vbExclamation
        Exit Sub
    End If

    originalCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo CleanFail
    gHasRequested = False

    ' 2. Clear or create FnGuide sheet
    Set wsFnGuide = GetOrCreateSheet("FnGuide")
    EnsureSheetWritable wsFnGuide
    ClearOutputSheet wsFnGuide

    startCol = 1
    companyCount = 0

    ' 3. Loop over companies
    For r = 3 To lastRow
        code = Trim$(CStr(wsMain.Cells(r, 2).Value))
        If Len(code) > 0 Then
            code = NormalizeStockCode(code)
            If Len(code) = 6 Then
                reportGbVal = Trim$(CStr(wsMain.Cells(r, 3).Value))
                reportGb = ResolveReportGb(reportGbVal, reportName)
                
                ' Add a 0.3s delay between companies
                If companyCount > 0 Then
                    Sleep 300
                End If
                
                companyCount = companyCount + 1
                gCompanyName = "" ' Clear before fetch
                
                ' Download and write side-by-side
                totalCols = DownloadFinanceCombinedForCol(wsFnGuide, code, reportGb, startCol)
                
                ' Move startCol to the right (leaving 1 blank column)
                startCol = startCol + totalCols + 1
            End If
        End If
    Next r

    ' 4. Final Formatting
    If companyCount > 0 Then
        wsFnGuide.Columns.AutoFit
    End If

    Application.Calculation = originalCalculation
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ShowMsgBox "FnGuide table download complete. Processed " & companyCount & " companies.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ShowMsgBox "Download failed: " & Err.Description & vbCrLf & "Step: " & gStep, vbCritical
End Sub

Private Function DownloadFinanceCombinedForCol(ByVal ws As Worksheet, ByVal code As String, ByVal reportGb As String, ByVal startCol As Long) As Long
    Dim fsTables As Collection
    Dim ratioTables As Collection

    gStep = "Load financial statements for A" & code
    Set fsTables = FetchDataTables(BuildFinanceUrl(code, reportGb), 6, "financial statement tables")

    gStep = "Load ratio tables for A" & code
    Set ratioTables = FetchDataTables(BuildRatioUrl(code, reportGb), 2, "ratio tables")

    ' If company name was not parsed or is empty, use a fallback
    If Len(gCompanyName) = 0 Then gCompanyName = "Code " & code

    gStep = "Write FnGuide for A" & code
    DownloadFinanceCombinedForCol = WriteSingleFinanceSheet(ws, fsTables, ratioTables, startCol, gCompanyName, "A" & code)
End Function

Private Sub DownloadFinanceCombined(ByVal code As String, ByVal reportGb As String)
    Dim fsTables As Collection
    Dim ratioTables As Collection

    gStep = "Load financial statements"
    Set fsTables = FetchDataTables(BuildFinanceUrl(code, reportGb), 6, "financial statement tables")

    gStep = "Load ratio tables"
    Set ratioTables = FetchDataTables(BuildRatioUrl(code, reportGb), 2, "ratio tables")

    gStep = "Write FS_Combined"
    WriteSingleFinanceSheet "FS_Combined", fsTables, ratioTables
End Sub

Private Sub DownloadRatioStacked(ByVal code As String, ByVal reportGb As String)
    Dim tables As Collection

    gStep = "Load ratio page"
    Set tables = FetchDataTables(BuildRatioUrl(code, reportGb), 2, "ratio tables")

    gStep = "Write RAT_Ratio"
    WriteStackedTablesSheet "RAT_Ratio", tables
End Sub

Private Sub WriteStackedTablesSheet(ByVal sheetName As String, ByVal tables As Collection)
    Dim ws As Worksheet
    Dim data As Variant
    Dim nextRow As Long
    Dim i As Long
    Dim rowCount As Long

    gStep = "Get sheet " & sheetName
    Set ws = GetOrCreateSheet(sheetName)
    EnsureSheetWritable ws
    ClearOutputSheet ws

    nextRow = 1
    For i = 1 To tables.Count
        gStep = "Parse table " & i & " for " & sheetName
        data = TableToArray(tables(i))
        rowCount = UBound(data, 1)

        gStep = "Write table " & i & " for " & sheetName
        ws.Cells(nextRow, 1).Resize(rowCount, UBound(data, 2)).Value = data

        gStep = "Format table " & i & " for " & sheetName
        ApplyTableFormattingAt ws, tables(i), rowCount, 1, nextRow - 1
        ws.Rows(nextRow).Font.Bold = True

        nextRow = nextRow + rowCount + 2
    Next i

    ws.Columns.AutoFit
End Sub

Private Function FetchDataTables(ByVal url As String, ByVal minTableCount As Long, ByVal pageName As String) As Collection
    Dim attempt As Long
    Dim html As String
    Dim tables As Collection

    For attempt = 1 To REQUEST_RETRY_COUNT
        gStep = "Fetch " & pageName & " (" & attempt & "/" & REQUEST_RETRY_COUNT & ")"
        html = FetchUtf8(url)
        Set tables = FindDataTables(html)

        If tables.Count >= minTableCount Then
            Set FetchDataTables = tables
            Exit Function
        End If

        If attempt < REQUEST_RETRY_COUNT Then Sleep REQUEST_RETRY_DELAY_MILLISECONDS
    Next attempt

    Err.Raise vbObjectError + 1000, , "Fewer " & pageName & " than expected. Expected " & _
        minTableCount & ", found " & tables.Count & ": " & url
End Function

Private Function FetchUtf8(ByVal url As String) As String
    Dim http As Object
    Dim stm As Object

    WaitBeforeRequest

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "Mozilla/5.0"
    http.SetRequestHeader "Referer", "https://comp.fnguide.com/"
    http.Send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1001, , "HTTP " & http.Status & ": " & url
    End If

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write http.ResponseBody
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "utf-8"
    FetchUtf8 = stm.ReadText
    stm.Close
    gHasRequested = True

    If InStr(1, url, "SVD_Finance.asp", vbTextCompare) > 0 Then
        gCompanyName = ExtractCompanyName(FetchUtf8)
    End If
End Function

Private Sub WaitBeforeRequest()
    If gHasRequested Then Sleep REQUEST_DELAY_MILLISECONDS
End Sub

Private Function RegExpDefault() As Object
    If gRegExpDefault Is Nothing Then Set gRegExpDefault = CreateObject("VBScript.RegExp")
    Set RegExpDefault = gRegExpDefault
End Function

Private Function RegExpRow() As Object
    If gRegExpRow Is Nothing Then Set gRegExpRow = CreateObject("VBScript.RegExp")
    Set RegExpRow = gRegExpRow
End Function

Private Function RegExpCell() As Object
    If gRegExpCell Is Nothing Then Set gRegExpCell = CreateObject("VBScript.RegExp")
    Set RegExpCell = gRegExpCell
End Function

Private Function RegExpSpan() As Object
    If gRegExpSpan Is Nothing Then Set gRegExpSpan = CreateObject("VBScript.RegExp")
    Set RegExpSpan = gRegExpSpan
End Function

Private Function ExtractBodyHtml(ByVal htmlText As String) As String
    Dim lowerText As String
    Dim p1 As Long
    Dim p2 As Long
    Dim p3 As Long

    lowerText = LCase$(htmlText)
    p1 = InStr(1, lowerText, "<body", vbBinaryCompare)

    If p1 > 0 Then
        p2 = InStr(p1, lowerText, ">", vbBinaryCompare)
        p3 = InStr(p2 + 1, lowerText, "</body>", vbBinaryCompare)
        If p2 > 0 And p3 > p2 Then
            ExtractBodyHtml = Mid$(htmlText, p2 + 1, p3 - p2 - 1)
            Exit Function
        End If
    End If

    ExtractBodyHtml = htmlText
End Function

Private Function FindDataTables(ByVal htmlText As String) As Collection
    Dim result As New Collection
    Dim m As Object
    Dim tableHtml As String
    Dim firstTag As String

    For Each m In RegexMatches(htmlText, "<table\b[\s\S]*?</table>", RegExpRow())
        tableHtml = CStr(m.Value)
        firstTag = Left$(tableHtml, InStr(1, tableHtml, ">", vbBinaryCompare))

        If InStr(1, firstTag, "us_table_ty1", vbTextCompare) > 0 _
           And InStr(1, firstTag, "zigbg_no", vbTextCompare) > 0 Then
            result.Add tableHtml
        End If
    Next m

    Set FindDataTables = result
End Function

Private Function RegexMatches(ByVal sourceText As String, ByVal pattern As String, Optional ByVal reObj As Object = Nothing) As Object
    Dim re As Object
    If reObj Is Nothing Then Set re = RegExpDefault() Else Set re = reObj

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = pattern
    Set RegexMatches = re.Execute(sourceText)
End Function

Private Sub WriteTableToSheet(ByVal tbl As String, ByVal sheetName As String)
    Dim ws As Worksheet
    Dim data As Variant

    gStep = "Get sheet " & sheetName
    Set ws = GetOrCreateSheet(sheetName)
    EnsureSheetWritable ws
    ClearOutputSheet ws

    gStep = "Parse table " & sheetName
    data = TableToArray(tbl)
    gStep = "Write values " & sheetName
    ws.Range("A1").Resize(UBound(data, 1), UBound(data, 2)).Value = data
    gStep = "Format " & sheetName
    ApplyTableFormatting ws, tbl, UBound(data, 1)
    ws.Columns.AutoFit
    ws.Rows(1).Font.Bold = True
End Sub

Private Function WriteSingleFinanceSheet( _
    ByVal ws As Worksheet, _
    ByVal fsTables As Collection, _
    ByVal ratioTables As Collection, _
    ByVal startCol As Long, _
    ByVal companyName As String, _
    ByVal stockCode As String) As Long

    Dim nextRow As Long
    Dim blockCols As Long
    Dim maxCols As Long

    maxCols = 0

    ' 1. Gather baseline headers from IS block
    Dim isAnnualData As Variant
    Dim isQuarterData As Variant
    isAnnualData = RemoveColumnsByHeader(TableToArray(fsTables(1)), Array(KoreanYoYPercentText()))
    isAnnualData = ProcessAnnualData(isAnnualData)
    
    isQuarterData = RemoveColumnsByHeader(TableToArray(fsTables(2)), Array(KoreanYoYPercentText()))
    isQuarterData = RemoveFirstColumn(isQuarterData)
    
    Dim annualHeaders As Variant
    Dim quarterHeaders As Variant
    Dim cols As Long, c As Long
    
    cols = UBound(isAnnualData, 2)
    ReDim annualHeaders(1 To cols - 1)
    For c = 2 To cols
        annualHeaders(c - 1) = isAnnualData(1, c)
    Next c
    
    cols = UBound(isQuarterData, 2)
    ReDim quarterHeaders(1 To cols)
    For c = 1 To cols
        quarterHeaders(c) = isQuarterData(1, c)
    Next c

    ' 2. Stack blocks consecutively
    nextRow = 1
    gStep = "Write IS to FnGuide"
    nextRow = WriteFinanceBlockStacked(ws, fsTables(1), fsTables(2), "IS", nextRow, True, Empty, Empty, startCol, companyName, stockCode, blockCols)
    If blockCols > maxCols Then maxCols = blockCols

    gStep = "Write BS to FnGuide"
    nextRow = WriteFinanceBlockStacked(ws, fsTables(3), fsTables(4), "BS", nextRow, False, annualHeaders, quarterHeaders, startCol, companyName, stockCode, blockCols)
    If blockCols > maxCols Then maxCols = blockCols

    gStep = "Write CFS to FnGuide"
    nextRow = WriteFinanceBlockStacked(ws, fsTables(5), fsTables(6), "CFS", nextRow, False, annualHeaders, quarterHeaders, startCol, companyName, stockCode, blockCols)
    If blockCols > maxCols Then maxCols = blockCols

    gStep = "Write Ratio to FnGuide"
    nextRow = WriteFinanceBlockStacked(ws, ratioTables(1), ratioTables(2), "Ratio", nextRow, False, annualHeaders, quarterHeaders, startCol, companyName, stockCode, blockCols)
    If blockCols > maxCols Then maxCols = blockCols

    ws.Columns(startCol).HorizontalAlignment = xlCenter
    
    ' Return total columns occupied (including classification column)
    WriteSingleFinanceSheet = maxCols + 1
End Function

Private Function GetFinanceTableColCount(ByVal tbl As String) As Long
    Dim data As Variant
    data = RemoveColumnsByHeader(TableToArray(tbl), Array(KoreanYoYPercentText()))
    GetFinanceTableColCount = UBound(data, 2)
End Function

Private Sub ClearOutputSheet(ByVal ws As Worksheet)
    gStep = "Clear sheet " & ws.Name

    If WorksheetFunction.CountA(ws.Cells) > 0 Then
        ws.UsedRange.Clear
    End If
End Sub

Private Sub EnsureSheetWritable(ByVal ws As Worksheet)
    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
        Err.Raise vbObjectError + 1005, , "Sheet is protected: " & ws.Name
    End If
End Sub

Private Function WriteFinanceBlockStacked( _
    ByVal ws As Worksheet, _
    ByVal annualTable As String, _
    ByVal quarterTable As String, _
    ByVal classFlag As String, _
    ByVal startRow As Long, _
    ByVal writeHeader As Boolean, _
    ByVal keepAnnualHeaders As Variant, _
    ByVal keepQuarterHeaders As Variant, _
    ByVal startCol As Long, _
    ByVal companyName As String, _
    ByVal stockCode As String, _
    ByRef outCols As Long) As Long

    Dim annualRows As Long
    Dim annualCols As Long
    Dim quarterRows As Long
    Dim quarterCols As Long
    Dim maxRows As Long
    Dim r As Long
    Dim quarterStartCol As Long

    annualRows = WriteFinanceBlockAt(ws, annualTable, startCol + 1, startRow, KoreanAnnualText(), writeHeader, keepAnnualHeaders, companyName, stockCode, annualCols)
    
    quarterStartCol = startCol + 1 + annualCols
    quarterRows = WriteFinanceBlockAt(ws, quarterTable, quarterStartCol, startRow, KoreanQuarterText(), writeHeader, keepQuarterHeaders, companyName, stockCode, quarterCols)

    If annualRows > quarterRows Then
        maxRows = annualRows
    Else
        maxRows = quarterRows
    End If

    ' Annual classification (Column startCol)
    If writeHeader Then
        ws.Cells(startRow, startCol).Value = Empty
        ws.Cells(startRow + 1, startCol).Value = Empty
        ws.Cells(startRow + 2, startCol).Value = KoreanClassificationText()
        For r = startRow + 3 To startRow + maxRows - 1
            ws.Cells(r, startCol).Value = classFlag
        Next r
        
        ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 2, startCol + 1 + annualCols + quarterCols - 1)).Font.Bold = True
    Else
        For r = startRow To startRow + maxRows - 1
            ws.Cells(r, startCol).Value = classFlag
        Next r
    End If

    outCols = annualCols + quarterCols

    WriteFinanceBlockStacked = startRow + maxRows
End Function

Private Function WriteFinanceBlockAt( _
    ByVal ws As Worksheet, _
    ByVal tbl As String, _
    ByVal startCol As Long, _
    ByVal startRow As Long, _
    ByVal periodFlag As String, _
    ByVal writeHeader As Boolean, _
    ByVal keepHeaders As Variant, _
    ByVal companyName As String, _
    ByVal stockCode As String, _
    ByRef outColsN As Long) As Long

    Dim data As Variant
    Dim dataOut As Variant
    Dim rowsN As Long
    Dim colsN As Long
    Dim blockName As String

    blockName = periodFlag

    gStep = "Parse block " & ws.Name & " / " & blockName
    data = GetProcessedData(tbl, periodFlag)
    
    If Not IsEmpty(keepHeaders) Then
        data = KeepOnlySpecificHeaders(data, keepHeaders)
    End If
    
    rowsN = UBound(data, 1)
    colsN = UBound(data, 2)
    outColsN = colsN

    If writeHeader Then
        ' Excel Row 1: periodFlag
        ' Excel Row 2: companyName or stockCode
        ' Excel Row 3: headers
        ' Excel Row 4+: data
        ReDim dataOut(1 To rowsN + 2, 1 To colsN)
        Dim r As Long, c As Long
        For c = 1 To colsN
            If periodFlag = KoreanAnnualText() And c = 1 Then
                dataOut(1, c) = Empty
                dataOut(2, c) = companyName
            Else
                dataOut(1, c) = periodFlag
                dataOut(2, c) = stockCode
            End If
            
            For r = 1 To rowsN
                dataOut(r + 2, c) = data(r, c)
            Next r
        Next c

        gStep = "Write block " & ws.Name & " / " & blockName
        ws.Cells(startRow, startCol).Resize(rowsN + 2, colsN).Value = dataOut
        
        If periodFlag = KoreanAnnualText() Then
            gStep = "Format block " & ws.Name & " / " & blockName
            ApplyTableFormattingAt ws, tbl, rowsN, startCol, startRow + 1, 0
        End If

        WriteFinanceBlockAt = rowsN + 2
    Else
        If rowsN > 1 Then
            ReDim dataOut(1 To rowsN - 1, 1 To colsN)
            Dim r2 As Long, c2 As Long
            For r2 = 2 To rowsN
                For c2 = 1 To colsN
                    dataOut(r2 - 1, c2) = data(r2, c2)
                Next c2
            Next r2

            gStep = "Write block (no header) " & ws.Name & " / " & blockName
            ws.Cells(startRow, startCol).Resize(rowsN - 1, colsN).Value = dataOut
            
            If periodFlag = KoreanAnnualText() Then
                gStep = "Format block " & ws.Name & " / " & blockName
                ApplyTableFormattingAt ws, tbl, rowsN - 1, startCol, startRow - 1, 1
            End If

            WriteFinanceBlockAt = rowsN - 1
        Else
            WriteFinanceBlockAt = 0
        End If
    End If
End Function

Private Sub FillFlagRow(ByRef data As Variant, ByVal rowIndex As Long, ByVal colCount As Long, ByVal value As String)
    Dim c As Long

    For c = 1 To colCount
        data(rowIndex, c) = value
    Next c
End Sub

Private Sub CopyArray(ByRef source As Variant, ByRef target As Variant, ByVal targetRowOffset As Long)
    Dim r As Long
    Dim c As Long

    For r = 1 To UBound(source, 1)
        For c = 1 To UBound(source, 2)
            target(r + targetRowOffset - 1, c) = source(r, c)
        Next c
    Next r
End Sub

Private Function RemoveColumnsByHeader(ByVal data As Variant, ByVal removeHeaders As Variant) As Variant
    Dim keep() As Boolean
    Dim c As Long
    Dim r As Long
    Dim outCol As Long
    Dim keepCount As Long
    Dim result() As Variant

    ReDim keep(1 To UBound(data, 2))

    For c = 1 To UBound(data, 2)
        keep(c) = Not IsInArray(CStr(data(1, c)), removeHeaders)
        If keep(c) Then keepCount = keepCount + 1
    Next c

    ReDim result(1 To UBound(data, 1), 1 To keepCount)

    For r = 1 To UBound(data, 1)
        outCol = 0
        For c = 1 To UBound(data, 2)
            If keep(c) Then
                outCol = outCol + 1
                result(r, outCol) = data(r, c)
            End If
        Next c
    Next r

    RemoveColumnsByHeader = result
End Function

Private Function IsInArray(ByVal value As String, ByVal items As Variant) As Boolean
    Dim i As Long

    For i = LBound(items) To UBound(items)
        If value = CStr(items(i)) Then
            IsInArray = True
            Exit Function
        End If
    Next i
End Function

Private Sub ApplyTableFormatting(ByVal ws As Worksheet, ByVal tbl As String, ByVal rowCount As Long)
    ApplyTableFormattingAt ws, tbl, rowCount, 1, 0, 0
End Sub

Private Sub ApplyTableFormattingAt( _
    ByVal ws As Worksheet, _
    ByVal tbl As String, _
    ByVal rowCount As Long, _
    ByVal startCol As Long, _
    ByVal rowOffset As Long, _
    ByVal trOffset As Long)

    Dim r As Long
    Dim targetRow As Long
    Dim className As String
    Dim m As Object
    Dim matches As Object

    If ws Is Nothing Then Exit Sub
    If rowCount <= 0 Then Exit Sub
    If startCol < 1 Or startCol > ws.Columns.Count Then Exit Sub

    Set matches = RegexMatches(tbl, "<tr\b([^>]*)>[\s\S]*?</tr>", RegExpRow())
    If matches Is Nothing Then Exit Sub

    r = 0
    Dim matchIdx As Long
    matchIdx = 0
    For Each m In matches
        matchIdx = matchIdx + 1
        If matchIdx <= trOffset Then GoTo NextMatch

        r = r + 1
        If r > rowCount Then Exit For

        targetRow = r + rowOffset
        If targetRow > ws.Rows.Count Then Exit For

        className = CStr(m.SubMatches(0))

        If targetRow >= 1 Then
            If InStr(1, className, "acd_dep2_sub", vbTextCompare) > 0 _
               Or InStr(1, className, "acd_dep_sub", vbTextCompare) > 0 Then
                ws.Cells(targetRow, startCol).IndentLevel = 1
            End If
        End If
NextMatch:
    Next m
End Sub

Private Function TableToArray(ByVal tbl As String) As Variant
    Dim rCount As Long
    Dim cCount As Long
    Dim data() As Variant
    Dim occupied() As Boolean
    Dim r As Long
    Dim c As Long
    Dim rowMatch As Object
    Dim cellMatch As Object
    Dim cellAttrs As String
    Dim rs As Long
    Dim cs As Long
    Dim rr As Long
    Dim cc As Long
    Dim rows As Object
    Dim cells As Object

    tbl = PreCleanTableHtml(tbl)

    Set rows = RegexMatches(tbl, "<tr\b[^>]*>[\s\S]*?</tr>", RegExpRow())
    rCount = rows.Count
    cCount = MaxColumnCount(tbl)

    ReDim data(1 To rCount, 1 To cCount)
    ReDim occupied(1 To rCount, 1 To cCount)

    r = 0
    For Each rowMatch In rows
        r = r + 1
        c = 1

        Set cells = RegexMatches(CStr(rowMatch.Value), "<t[hd]\b([^>]*)>([\s\S]*?)</t[hd]>", RegExpCell())
        For Each cellMatch In cells
            Do While c <= cCount And occupied(r, c)
                c = c + 1
            Loop

            If c <= cCount Then
                cellAttrs = CStr(cellMatch.SubMatches(0))
                data(r, c) = HtmlToText(CStr(cellMatch.SubMatches(1)))

                rs = HtmlSpan(cellAttrs, "rowspan")
                cs = HtmlSpan(cellAttrs, "colspan")

                For rr = r To MinLong(r + rs - 1, rCount)
                    For cc = c To MinLong(c + cs - 1, cCount)
                        occupied(rr, cc) = True
                    Next cc
                Next rr

                c = c + cs
            End If
        Next cellMatch
    Next rowMatch

    TableToArray = data
End Function

Private Function MaxColumnCount(ByVal tbl As String) As Long
    Dim rowMatch As Object
    Dim cellMatch As Object
    Dim n As Long
    Dim maxN As Long

    For Each rowMatch In RegexMatches(tbl, "<tr\b[^>]*>[\s\S]*?</tr>", RegExpRow())
        n = 0
        For Each cellMatch In RegexMatches(CStr(rowMatch.Value), "<t[hd]\b([^>]*)>[\s\S]*?</t[hd]>", RegExpCell())
            n = n + HtmlSpan(CStr(cellMatch.SubMatches(0)), "colspan")
        Next cellMatch
        If n > maxN Then maxN = n
    Next rowMatch

    MaxColumnCount = maxN
End Function

Private Function HtmlSpan(ByVal attrs As String, ByVal attrName As String) As Long
    Dim matches As Object

    Set matches = RegexMatches(attrs, attrName & "\s*=\s*[""']?([0-9]+)", RegExpSpan())
    If matches.Count = 0 Then
        HtmlSpan = 1
    Else
        HtmlSpan = CLng(matches(0).SubMatches(0))
        If HtmlSpan < 1 Then HtmlSpan = 1
    End If
End Function

Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then
        MinLong = a
    Else
        MinLong = b
    End If
End Function

Private Function PreCleanTableHtml(ByVal tbl As String) As String
    Dim s As String
    s = RegexReplace(tbl, "<script\b[\s\S]*?</script>", " ", RegExpDefault())
    s = RegexReplace(s, "<style\b[\s\S]*?</style>", " ", RegExpDefault())
    s = RegexReplace(s, "<dl\b[\s\S]*?</dl>", " ", RegExpDefault())
    s = RegexReplace(s, "<a\b[^>]*\bbtn_acdopen\b[^>]*>[\s\S]*?</a>", " ", RegExpDefault())
    s = RegexReplace(s, "<a\b[^>]*\bbtn_acdclose\b[^>]*>[\s\S]*?</a>", " ", RegExpDefault())
    PreCleanTableHtml = s
End Function

Private Function HtmlToText(ByVal htmlText As String) As String
    Dim s As String

    s = RegexReplace(htmlText, "<[^>]+>", " ", RegExpDefault())
    s = HtmlDecodeBasic(s)
    HtmlToText = CleanCellText(s)
End Function

Private Function RegexReplace(ByVal sourceText As String, ByVal pattern As String, ByVal replacement As String, Optional ByVal reObj As Object = Nothing) As String
    Dim re As Object
    If reObj Is Nothing Then Set re = RegExpDefault() Else Set re = reObj

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = pattern
    RegexReplace = re.Replace(sourceText, replacement)
End Function

Private Function HtmlDecodeBasic(ByVal s As String) As String
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&#160;", " ")
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&#39;", "'")
    HtmlDecodeBasic = s
End Function

Private Function CleanCellText(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, KoreanExpandText(), "")
    s = Replace(s, KoreanCloseText(), "")
    s = Replace(s, KoreanExpandText2(), "")
    s = Replace(s, KoreanCloseText2(), "")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanCellText = Trim$(s)
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim wb As Workbook

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = ThisWorkbook

    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function NormalizeStockCode(ByVal code As String) As String
    Dim i As Long
    Dim ch As String
    Dim digits As String

    code = Replace(code, "A", "", 1, -1, vbTextCompare)
    code = Trim$(code)

    For i = 1 To Len(code)
        ch = Mid$(code, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i

    If Len(digits) = 0 Or Len(digits) > 6 Then
        NormalizeStockCode = ""
    Else
        NormalizeStockCode = Right$("000000" & digits, 6)
    End If
End Function

Private Function GetMainValue(ByVal address As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets("MAIN")
    On Error GoTo 0

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1002, , "MAIN sheet was not found. Create MAIN sheet and enter C2/C3."
    End If

    GetMainValue = Trim$(CStr(ws.Range(address).Value))
End Function

Private Function ResolveReportGb(ByVal value As String, ByRef reportName As String) As String
    value = Trim$(value)

    Select Case UCase$(value)
        Case "", "CON", "CONS", "CONSOLIDATED", "A"
            reportName = "Consolidated"
            ResolveReportGb = ""
        Case "B", "SEP", "SEPARATE", "SEPARATED", "STANDALONE"
            reportName = "Separate"
            ResolveReportGb = "B"
        Case Else
            If value = KoreanConsolidated() Then
                reportName = "Consolidated"
                ResolveReportGb = ""
            ElseIf value = KoreanSeparate() Then
                reportName = "Separate"
                ResolveReportGb = "B"
            Else
                Err.Raise vbObjectError + 1003, , "MAIN!C3 must be Consolidated/Separate or Korean equivalent."
            End If
    End Select
End Function

Private Function KoreanConsolidated() As String
    KoreanConsolidated = ChrW(&HC5F0) & ChrW(&HACB0)
End Function

Private Function KoreanSeparate() As String
    KoreanSeparate = ChrW(&HBCC4) & ChrW(&HB3C4)
End Function

Private Function KoreanAnnualText() As String
    KoreanAnnualText = ChrW(&HC5F0) & ChrW(&HAC04)
End Function

Private Function KoreanQuarterText() As String
    KoreanQuarterText = ChrW(&HBD84) & ChrW(&HAE30)
End Function

Private Function KoreanYoYPercentText() As String
    KoreanYoYPercentText = ChrW(&HC804) & ChrW(&HB144) & _
                           ChrW(&HB3D9) & ChrW(&HAE30) & "(%)"
End Function

Private Function KoreanExpandText() As String
    KoreanExpandText = ChrW(&HACC4) & ChrW(&HC0B0) & ChrW(&HC5D0) & " " & _
                       ChrW(&HCC38) & ChrW(&HC5EC) & ChrW(&HD55C) & " " & _
                       ChrW(&HACC4) & ChrW(&HC815) & " " & _
                       ChrW(&HD3BC) & ChrW(&HCE58) & ChrW(&HAE30)
End Function

Private Function KoreanCloseText() As String
    KoreanCloseText = ChrW(&HACC4) & ChrW(&HC0B0) & ChrW(&HC5D0) & " " & _
                      ChrW(&HCC38) & ChrW(&HC5EC) & ChrW(&HD55C) & " " & _
                      ChrW(&HACC4) & ChrW(&HC815) & " " & _
                      ChrW(&HB2EB) & ChrW(&HAE30)
End Function

Private Function KoreanExpandText2() As String
    KoreanExpandText2 = ChrW(&HACC4) & ChrW(&HC0B0) & " " & _
                        ChrW(&HCC38) & ChrW(&HC5EC) & ChrW(&HD55C) & " " & _
                        ChrW(&HACC4) & ChrW(&HC815) & " " & _
                        ChrW(&HD3BC) & ChrW(&HCE58) & ChrW(&HAE30)
End Function

Private Function KoreanCloseText2() As String
    KoreanCloseText2 = ChrW(&HACC4) & ChrW(&HC0B0) & " " & _
                       ChrW(&HCC38) & ChrW(&HC5EC) & ChrW(&HD55C) & " " & _
                       ChrW(&HACC4) & ChrW(&HC815) & " " & _
                       ChrW(&HB2EB) & ChrW(&HAE30)
End Function

Private Function BuildFinanceUrl(ByVal code As String, ByVal reportGb As String) As String
    BuildFinanceUrl = BASE_URL & "SVD_Finance.asp?pGB=1&gicode=A" & code & _
                      "&cID=AA&MenuYn=Y&ReportGB=" & reportGb & "&NewMenuID=103&stkGb=701"
End Function

Private Function BuildRatioUrl(ByVal code As String, ByVal reportGb As String) As String
    BuildRatioUrl = BASE_URL & "SVD_FinanceRatio.asp?pGB=1&gicode=A" & code & _
                    "&cID=AA&MenuYn=Y&ReportGB=" & reportGb & "&NewMenuID=104&stkGb=701"
End Function

Private Function KoreanClassificationText() As String
    KoreanClassificationText = ChrW(&HAD6C) & ChrW(&HBD84)
End Function

Private Function KoreanYoYText() As String
    KoreanYoYText = ChrW(&HC804) & ChrW(&HB144) & ChrW(&HB3D9) & ChrW(&HAE30)
End Function

Private Function ProcessAnnualData(ByVal data As Variant) As Variant
    Dim c As Long, r As Long
    Dim colsN As Long, rowsN As Long
    colsN = UBound(data, 2)
    rowsN = UBound(data, 1)

    ' 1. Collect columns with YYYY/MM pattern
    Dim ymCols() As Long
    Dim ymMonths() As String
    Dim ymCount As Long
    ymCount = 0

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\b\d{4}/(\d{2})\b"
    re.Global = False
    re.IgnoreCase = True

    For c = 2 To colsN
        Dim header As String
        header = CStr(data(1, c))
        If re.Test(header) Then
            Dim matches As Object
            Set matches = re.Execute(header)
            ymCount = ymCount + 1
            ReDim Preserve ymCols(1 To ymCount)
            ReDim Preserve ymMonths(1 To ymCount)
            ymCols(ymCount) = c
            ymMonths(ymCount) = matches(0).SubMatches(0) ' Extract MM
        End If
    Next c

    ' 2. Check condition: if the rightmost YYYY/MM month is unique
    Dim shouldRemove As Boolean
    shouldRemove = False
    If ymCount >= 2 Then
        Dim lastMonth As String
        Dim firstMonth As String
        lastMonth = ymMonths(ymCount)
        firstMonth = ymMonths(1)
        
        Dim allPrevSame As Boolean
        allPrevSame = True
        Dim i As Long
        For i = 1 To ymCount - 1
            If ymMonths(i) <> firstMonth Then
                allPrevSame = False
                Exit For
            End If
        Next i
        
        If allPrevSame And lastMonth <> firstMonth Then
            shouldRemove = True
        End If
    End If

    ' 3. Process removal
    If shouldRemove Then
        Dim removeColLastYM As Long
        removeColLastYM = ymCols(ymCount)
        
        Dim removeColYoY As Long
        removeColYoY = 0
        Dim yoyText As String
        yoyText = KoreanYoYText()
        
        For c = 2 To colsN
            If CStr(data(1, c)) = yoyText Then
                removeColYoY = c
                Exit For
            End If
        Next c
        
        Dim keepCount As Long
        keepCount = 0
        Dim keepCols() As Long
        For c = 1 To colsN
            If c <> removeColLastYM And c <> removeColYoY Then
                keepCount = keepCount + 1
                ReDim Preserve keepCols(1 To keepCount)
                keepCols(keepCount) = c
            End If
        Next c
        
        Dim result() As Variant
        ReDim result(1 To rowsN, 1 To keepCount)
        For r = 1 To rowsN
            For c = 1 To keepCount
                result(r, c) = data(r, keepCols(c))
            Next c
        Next r
        ProcessAnnualData = result
    Else
        ProcessAnnualData = data
    End If
End Function

Private Function RemoveFirstColumn(ByVal data As Variant) As Variant
    Dim r As Long, c As Long
    Dim rowsN As Long, colsN As Long
    rowsN = UBound(data, 1)
    colsN = UBound(data, 2)
    
    If colsN <= 1 Then
        RemoveFirstColumn = data
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To rowsN, 1 To colsN - 1)
    
    For r = 1 To rowsN
        For c = 2 To colsN
            result(r, c - 1) = data(r, c)
        Next c
    Next r
    
    RemoveFirstColumn = result
End Function

Private Function GetCleanHeaderKey(ByVal header As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\b\d{4}/\d{2}\b"
    re.Global = False
    re.IgnoreCase = True
    If re.Test(header) Then
        GetCleanHeaderKey = re.Execute(header)(0).Value
    Else
        GetCleanHeaderKey = Trim$(header)
    End If
End Function

Private Function KeepOnlySpecificHeaders(ByVal data As Variant, ByVal keepHeaders As Variant) As Variant
    Dim keep() As Boolean
    Dim c As Long, r As Long
    Dim outCol As Long
    Dim keepCount As Long
    Dim result() As Variant
    Dim colsN As Long
    
    colsN = UBound(data, 2)
    ReDim keep(1 To colsN)
    
    Dim isFirstColAccountName As Boolean
    isFirstColAccountName = True
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\b\d{4}/\d{2}\b"
    If re.Test(CStr(data(1, 1))) Then
        isFirstColAccountName = False
    End If

    If isFirstColAccountName Then
        keep(1) = True
        keepCount = 1
    End If
    
    Dim startIdx As Long
    If isFirstColAccountName Then startIdx = 2 Else startIdx = 1
    
    For c = startIdx To colsN
        Dim headerVal As String
        headerVal = CStr(data(1, c))
        Dim cleanHeader As String
        cleanHeader = GetCleanHeaderKey(headerVal)
        
        Dim isFound As Boolean
        isFound = False
        Dim i As Long
        For i = LBound(keepHeaders) To UBound(keepHeaders)
            If GetCleanHeaderKey(CStr(keepHeaders(i))) = cleanHeader Then
                isFound = True
                Exit For
            End If
        Next i
        
        keep(c) = isFound
        If keep(c) Then keepCount = keepCount + 1
    Next c
    
    If keepCount = 0 Then
        ReDim result(1 To UBound(data, 1), 1 To 1)
        KeepOnlySpecificHeaders = result
        Exit Function
    End If
    
    ReDim result(1 To UBound(data, 1), 1 To keepCount)
    For r = 1 To UBound(data, 1)
        outCol = 0
        For c = 1 To colsN
            If keep(c) Then
                outCol = outCol + 1
                result(r, outCol) = data(r, c)
            End If
        Next c
    Next r
    
    KeepOnlySpecificHeaders = result
End Function

Private Function GetProcessedData(ByVal tbl As String, ByVal periodFlag As String) As Variant
    Dim data As Variant
    data = RemoveColumnsByHeader(TableToArray(tbl), Array(KoreanYoYPercentText()))
    
    If periodFlag = KoreanAnnualText() Then
        data = ProcessAnnualData(data)
    ElseIf periodFlag = KoreanQuarterText() Then
        data = RemoveFirstColumn(data)
    End If
    
    GetProcessedData = data
End Function

Private Function ExtractCompanyName(ByVal htmlText As String) As String
    Dim re As Object
    Dim matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False
    
    re.Pattern = "<h1 id=""giName"">([^<]+)</h1>"
    If re.Test(htmlText) Then
        Set matches = re.Execute(htmlText)
        ExtractCompanyName = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    re.Pattern = "id=""giname"" value=""([^""]+)"""
    If re.Test(htmlText) Then
        Set matches = re.Execute(htmlText)
        ExtractCompanyName = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    re.Pattern = "<title>([^(|]+)\("
    If re.Test(htmlText) Then
        Set matches = re.Execute(htmlText)
        ExtractCompanyName = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ExtractCompanyName = "Unknown"
End Function

Private Sub ShowMsgBox(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "")
    If Application.UserControl Then
        MsgBox prompt, buttons, title
    Else
        Debug.Print prompt
    End If
End Sub
