Attribute VB_Name = "OpenDartDownloader"
Option Explicit

' OpenDART Financial Statement Downloader (Optimized Version)
' Downloads financial statements from OpenDART in batches to maximize performance.
' Reads and writes worksheets using 2D Variant Arrays to avoid COM overhead.
' Caches company mappings in the DART corp-code sheet to run instantly after first load.
' Caches fetched financial statement records in-memory during execution.

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const BASE_API_URL As String = "https://opendart.fss.or.kr/api/"

Private Function WText(ParamArray codePoints() As Variant) As String
    Dim i As Long
    Dim s As String
    For i = LBound(codePoints) To UBound(codePoints)
        s = s & ChrW$(CLng(codePoints(i)))
    Next i
    WText = s
End Function

Private Function DartCorpCodeSheetName() As String
    DartCorpCodeSheetName = "DART " & WText(&HB0B4, &HBD80, &HCF54, &HB4DC)
End Function

Private Function DartCorpCodeInitLabel() As String
    DartCorpCodeInitLabel = DartCorpCodeSheetName() & WText(&H0020, &HCD08, &HAE30, &HD654)
End Function

Private Function DartCorpCodeInitDoneMessage(ByVal matchCount As Long) As String
    DartCorpCodeInitDoneMessage = DartCorpCodeInitLabel() & _
        WText(&H0020, &HC644, &HB8CC, &H002E, &H0020, &HCD1D, &H0020) & _
        matchCount & _
        WText(&HAC1C, &HC758, &H0020, &HD68C, &HC0AC, &H0020, &HC815, &HBCF4, &HAC00, &H0020, &HC0DD, &HC131, &HB418, &HC5C8, &HC2B5, &HB2C8, &HB2E4, &H002E)
End Function

Private Function DartCorpCodeMissingMessage() As String
    DartCorpCodeMissingMessage = DartCorpCodeSheetName() & _
        WText(&H0020, &HC2DC, &HD2B8, &HAC00, &H0020, &HC874, &HC7AC, &HD558, &HC9C0, &H0020, &HC54A, &HC2B5, &HB2C8, &HB2E4, &H002E, &H0020, &HBA3C, &HC800, &H0020, &H0027) & _
        DartCorpCodeInitLabel() & _
        WText(&H0027, &H0020, &HB9E4, &HD06C, &HB85C, &HB97C, &H0020, &HC2E4, &HD589, &HD574, &H0020, &HC8FC, &HC138, &HC694, &H002E)
End Function

Private Function DartCorpCodeEmptyMessage() As String
    DartCorpCodeEmptyMessage = DartCorpCodeSheetName() & _
        WText(&H0020, &HC2DC, &HD2B8, &HAC00, &H0020, &HBE44, &HC5B4, &H0020, &HC788, &HC2B5, &HB2C8, &HB2E4, &H002E, &H0020, &HBA3C, &HC800, &H0020, &H0027) & _
        DartCorpCodeInitLabel() & _
        WText(&H0027, &H0020, &HB9E4, &HD06C, &HB85C, &HB97C, &H0020, &HC2E4, &HD589, &HD574, &H0020, &HC8FC, &HC138, &HC694, &H002E)
End Function

Private Function DartCorpCodeMissingDescription() As String
    DartCorpCodeMissingDescription = DartCorpCodeSheetName() & WText(&H0020, &HC2DC, &HD2B8, &H0020, &HC5C6, &HC74C)
End Function

Private Function DartCorpCodeEmptyDescription() As String
    DartCorpCodeEmptyDescription = DartCorpCodeSheetName() & WText(&H0020, &HC2DC, &HD2B8, &H0020, &HBE44, &HC5B4, &HC788, &HC74C)
End Function

Private Function KoreanAll() As String
    KoreanAll = WText(&HC804, &HCCB4)
End Function

Private Function KoreanMajor() As String
    KoreanMajor = WText(&HC8FC, &HC694)
End Function

Private Function KoreanSummary() As String
    KoreanSummary = WText(&HC694, &HC57D)
End Function

Private Function KoreanYear() As String
    KoreanYear = WText(&HB144)
End Function

Private Function KoreanAnnual() As String
    KoreanAnnual = WText(&HC5F0, &HAC04)
End Function

Private Function KoreanSeparate() As String
    KoreanSeparate = WText(&HBCC4, &HB3C4)
End Function

Private Function KoreanConsolidated() As String
    KoreanConsolidated = WText(&HC5F0, &HACB0)
End Function

Private Function KoreanCompanyName() As String
    KoreanCompanyName = WText(&HAE30, &HC5C5, &HBA85)
End Function

Private Function KoreanStockCode() As String
    KoreanStockCode = WText(&HC885, &HBAA9, &HCF54, &HB4DC)
End Function

Private Function KoreanCategory() As String
    KoreanCategory = WText(&HAD6C, &HBD84)
End Function

Private Function KoreanCfsStatement() As String
    KoreanCfsStatement = WText(&HC5F0, &HACB0, &HC7AC, &HBB34, &HC81C, &HD45C)
End Function

Private Function KoreanOfsStatement() As String
    KoreanOfsStatement = WText(&HBCC4, &HB3C4, &HC7AC, &HBB34, &HC81C, &HD45C)
End Function

Private Function KoreanAccountName() As String
    KoreanAccountName = WText(&HACC4, &HC815, &HACFC, &HBAA9)
End Function

Private Function KoreanQuarter1() As String
    KoreanQuarter1 = WText(&H0031, &HBD84, &HAE30)
End Function

Private Function KoreanHalf() As String
    KoreanHalf = WText(&HBC18, &HAE30)
End Function

Private Function KoreanQuarter2() As String
    KoreanQuarter2 = WText(&H0032, &HBD84, &HAE30)
End Function

Private Function KoreanQuarter3() As String
    KoreanQuarter3 = WText(&H0033, &HBD84, &HAE30)
End Function

Private Function KoreanQuarter4() As String
    KoreanQuarter4 = WText(&H0034, &HBD84, &HAE30)
End Function

Private Function KoreanBusiness() As String
    KoreanBusiness = WText(&HC0AC, &HC5C5)
End Function

Private Function KoreanUnregisteredKey() As String
    KoreanUnregisteredKey = WText(&HBBF8, &HB4F1, &HB85D, &H0020, &HC778, &HC99D, &HD0A4)
End Function

Private Function KoreanInvalidKey() As String
    KoreanInvalidKey = WText(&HC0AC, &HC6A9, &HD560, &H0020, &HC218, &H0020, &HC5C6, &HB294, &H0020, &HC778, &HC99D, &HD0A4)
End Function

Private Function KoreanOtherError() As String
    KoreanOtherError = WText(&HAE30, &HD0C0, &H0020, &HC5D0, &HB7EC)
End Function

Private Function KoreanTargetCompany() As String
    KoreanTargetCompany = WText(&HB300, &HC0C1, &H0020, &HAE30, &HC5C5, &H003A, &H0020)
End Function

Public Sub TestLogging()
    LogMsg "Test logging from VBA"
End Sub

Public Sub TestAdodb()
    LogMsg "TestAdodb: Started"
    Dim cachedXmlPath As String
    cachedXmlPath = Environ("TEMP") & "\CORPCODE.xml"

    LogMsg "TestAdodb: Opening stream"
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile cachedXmlPath
    LogMsg "TestAdodb: Loaded from file"

    Dim xmlContent As String
    xmlContent = stream.ReadText(-1)
    LogMsg "TestAdodb: Read text, size = " & Len(xmlContent)

    stream.Close
    LogMsg "TestAdodb: Finished"
End Sub

Public Sub DownloadDartData()
    Dim wsMain As Worksheet
    Dim wsOut As Worksheet
    Dim apiKey As String
    Dim targetDateStr As String
    Dim accountType As String
    Dim isMajor As Boolean
    Dim wb As Workbook
    Dim lastRow As Long
    Dim r As Long
    Dim stockCode As String
    Dim corpCode As String
    Dim fsDivStr As String
    Dim fsDiv As String ' "CFS", "OFS", or "AUTO"

    Dim originalCalculation As Long
    Dim compKey As Variant
    Dim compInfo As Variant
    Dim cKey As Variant
    Dim parts() As String

    ' 1. Initialize Workbook and Sheet references
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = ThisWorkbook

    LogMsg "DownloadDartData: Macro execution started"

    On Error Resume Next
    Set wsMain = wb.Worksheets("MAIN")
    On Error GoTo 0

    If wsMain Is Nothing Then
        MsgBox "MAIN sheet was not found. Please create a sheet named 'MAIN'.", vbCritical
        Exit Sub
    End If

    ' 2. Read parameters
    apiKey = Trim$(CStr(wsMain.Range("C2").Value))
    targetDateStr = Trim$(CStr(wsMain.Range("C3").Value))
    accountType = Trim$(CStr(wsMain.Range("C4").Value))

    If Len(accountType) = 0 Then accountType = KoreanAll()
    isMajor = (LCase$(accountType) = KoreanMajor() Or LCase$(accountType) = KoreanSummary() Or LCase$(accountType) = "major" Or LCase$(accountType) = "summary")

    If Len(apiKey) = 0 Then
        MsgBox "API Key in cell C2 of MAIN sheet is empty.", vbCritical
        Exit Sub
    End If

    If Len(targetDateStr) = 0 Then
        MsgBox "Target date/quarter in cell C3 of MAIN sheet is empty.", vbCritical
        Exit Sub
    End If

    ' Parse target date
    Dim targetYear As Integer
    Dim targetReprtCode As String
    Dim targetPeriodName As String

    On Error GoTo ParseError
    ParseTargetDate targetDateStr, targetYear, targetReprtCode, targetPeriodName
    On Error GoTo CleanFail

    ' Determine periods
    Dim isAnnual As Boolean
    isAnnual = (targetReprtCode = "11011")

    Dim numPeriods As Integer
    Dim periodLabels() As String
    Dim periodYears() As Integer
    Dim periodRepCodes() As String

    If isAnnual Then
        numPeriods = 3
        ReDim periodLabels(1 To 3)
        ReDim periodYears(1 To 3)
        ReDim periodRepCodes(1 To 3)

        periodLabels(1) = CStr(targetYear - 2) & KoreanYear() & " " & KoreanAnnual()
        periodLabels(2) = CStr(targetYear - 1) & KoreanYear() & " " & KoreanAnnual()
        periodLabels(3) = CStr(targetYear) & KoreanYear() & " " & KoreanAnnual()

        periodYears(1) = targetYear - 2
        periodYears(2) = targetYear - 1
        periodYears(3) = targetYear

        periodRepCodes(1) = "11011"
        periodRepCodes(2) = "11011"
        periodRepCodes(3) = "11011"
    Else
        numPeriods = 5
        ReDim periodLabels(1 To 5)
        ReDim periodYears(1 To 5)
        ReDim periodRepCodes(1 To 5)

        periodLabels(1) = CStr(targetYear - 3) & KoreanYear() & " " & KoreanAnnual()
        periodLabels(2) = CStr(targetYear - 2) & KoreanYear() & " " & KoreanAnnual()
        periodLabels(3) = CStr(targetYear - 1) & KoreanYear() & " " & KoreanAnnual()
        periodLabels(4) = CStr(targetYear - 1) & KoreanYear() & " " & targetPeriodName
        periodLabels(5) = CStr(targetYear) & KoreanYear() & " " & targetPeriodName

        periodYears(1) = targetYear - 3
        periodYears(2) = targetYear - 2
        periodYears(3) = targetYear - 1
        periodYears(4) = targetYear - 1
        periodYears(5) = targetYear

        periodRepCodes(1) = "11011"
        periodRepCodes(2) = "11011"
        periodRepCodes(3) = "11011"
        periodRepCodes(4) = targetReprtCode
        periodRepCodes(5) = targetReprtCode
    End If

    ' Get target stock codes
    lastRow = wsMain.Cells(wsMain.Rows.Count, 2).End(xlUp).Row
    If lastRow < 6 Then
        MsgBox "No stock codes found in MAIN column B starting from row 6.", vbExclamation
        Exit Sub
    End If

    ' 3. Initialize mapping (Stock Code to Corp Code and Name)
    Dim corpMap As Object
    Dim nameMap As Object
    Set corpMap = CreateObject("Scripting.Dictionary")
    Set nameMap = CreateObject("Scripting.Dictionary")

    LogMsg "DownloadDartData: Starting mapping initialization"
    LoadCorpCodeMap corpMap, nameMap
    LogMsg "DownloadDartData: Mapping initialization complete. Companies in map: " & corpMap.Count

    ' Save application states
    originalCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 4. Initialize Cache Dict (In-Memory cache for deduplication during current run)
    Dim cacheDict As Object
    Set cacheDict = CreateObject("Scripting.Dictionary")

    ' 5. Clear/Create Output Sheet
    On Error Resume Next
    Set wsOut = wb.Worksheets("OpenDART")
    If Not wsOut Is Nothing Then
        wsOut.UsedRange.Clear
    Else
        Set wsOut = wb.Worksheets.Add(After:=wsMain)
        wsOut.Name = "OpenDART"
    End If
    On Error GoTo CleanFail

    ' Dictionaries to store company account lists and active company metadata
    Dim corpAccounts As Object
    Dim activeComps As Object
    Set corpAccounts = CreateObject("Scripting.Dictionary")
    Set activeComps = CreateObject("Scripting.Dictionary")

    ' 6. Collect active companies and resolve their fsDiv (Fast Range Read)
    Dim compIdx As Long
    compIdx = 0

    Dim mainRangeData As Variant
    mainRangeData = wsMain.Range(wsMain.Cells(6, 2), wsMain.Cells(lastRow, 3)).Value

    Dim idxRow As Long
    For idxRow = 1 To UBound(mainRangeData, 1)
        stockCode = Trim$(CStr(mainRangeData(idxRow, 1)))
        If Len(stockCode) > 0 Then
            stockCode = NormalizeStockCode(stockCode)
            If corpMap.Exists(stockCode) Then
                corpCode = corpMap(stockCode)
                fsDivStr = Trim$(CStr(mainRangeData(idxRow, 2)))

                ' Resolve consolidated vs separate. Blank means try consolidated first, then separate on 013.
                If fsDivStr = KoreanSeparate() Or UCase$(fsDivStr) = "OFS" Or UCase$(fsDivStr) = "SEP" Or UCase$(fsDivStr) = "SEPARATE" Then
                    fsDiv = "OFS"
                ElseIf fsDivStr = KoreanConsolidated() Or UCase$(fsDivStr) = "CFS" Or UCase$(fsDivStr) = "CONSOLIDATED" Then
                    fsDiv = "CFS"
                Else
                    fsDiv = "AUTO"
                End If

                activeComps(corpCode) = Array(stockCode, nameMap(corpCode), fsDiv)
                Set corpAccounts(corpCode) = New Collection
                compIdx = compIdx + 1
            End If
        End If
    Next idxRow

    ' 7. Batch fetch period data for all periods (No Sleep Delay)
    Dim p As Integer
    For p = 1 To numPeriods
        Dim fetchList As Collection
        Set fetchList = New Collection

        For Each compKey In activeComps.Keys
            compInfo = activeComps(compKey)
            fsDiv = compInfo(2)

            Dim statusKeyNew As String, statusKeyOld As String
            statusKeyNew = compKey & "|" & fsDiv & "|" & CStr(periodYears(p)) & "|" & periodRepCodes(p)
            statusKeyOld = compKey & "|" & CStr(periodYears(p)) & "|" & periodRepCodes(p)

            Dim isCached As Boolean
            isCached = False
            If cacheDict.Exists(statusKeyNew) Then
                isCached = (cacheDict(statusKeyNew) = "Fetched")
            ElseIf cacheDict.Exists(statusKeyOld) Then
                isCached = (cacheDict(statusKeyOld) = "Fetched")
            End If

            If Not isCached Then
                fetchList.Add compKey
            End If
        Next compKey

        If fetchList.Count > 0 Then
            ' Fallback: no data for this period - skip and leave blank
            On Error Resume Next
            FetchPeriodDataBatch apiKey, fetchList, periodYears(p), periodRepCodes(p), cacheDict, isMajor, activeComps
            If Err.Number <> 0 Then
                LogMsg "Period fallback: year=" & periodYears(p) & " reprt=" & periodRepCodes(p) & " skipped. Error=" & Err.Description
                Err.Clear
            End If
            On Error GoTo CleanFail
        End If
    Next p

    ' 8. Resolve AUTO companies to a single actual statement type before output.
    For Each compKey In activeComps.Keys
        compInfo = activeComps(compKey)
        If compInfo(2) = "AUTO" Then
            activeComps(compKey) = Array(compInfo(0), compInfo(1), ResolveAutoFsDiv(cacheDict, CStr(compKey)))
        End If
    Next compKey

    ' 9. Populate unique account lists from cacheDict (Supports new and old format)
    ' Loop over periods in reverse (latest first) to get accounts in the order of the latest report!
    For Each compKey In activeComps.Keys
        compInfo = activeComps(compKey)
        fsDiv = compInfo(2)

        Dim pRev As Integer
        Dim targetYr As Integer
        Dim targetRepCode As String

        Dim stmtIdx As Integer
        Dim targetSjDiv As String
        For stmtIdx = 1 To 5
            targetSjDiv = StatementDivByOrder(stmtIdx)
            For pRev = numPeriods To 1 Step -1
                targetYr = periodYears(pRev)
                targetRepCode = periodRepCodes(pRev)

                For Each cKey In cacheDict.Keys
                    parts = Split(cKey, "|")
                    If UBound(parts) = 5 Then ' New format: corpCode|fsDiv|sjDiv|accName|yr|repCode
                        If parts(0) = compKey And FsDivMatches(parts(1), fsDiv) Then
                            If parts(2) = targetSjDiv And CInt(parts(4)) = targetYr And parts(5) = targetRepCode Then
                                AddUnique corpAccounts(compKey), parts(2) & " | " & parts(3)
                            End If
                        End If
                    ElseIf UBound(parts) = 4 Then ' Old format: corpCode|sjDiv|accName|yr|repCode
                        If parts(0) = compKey Then
                            If parts(1) = targetSjDiv And CInt(parts(3)) = targetYr And parts(4) = targetRepCode Then
                                AddUnique corpAccounts(compKey), parts(1) & " | " & parts(2)
                            End If
                        End If
                    End If
                Next cKey
            Next pRev
        Next stmtIdx

        For pRev = numPeriods To 1 Step -1
            targetYr = periodYears(pRev)
            targetRepCode = periodRepCodes(pRev)

            For Each cKey In cacheDict.Keys
                parts = Split(cKey, "|")
                If UBound(parts) = 5 Then
                    If parts(0) = compKey And FsDivMatches(parts(1), fsDiv) Then
                        If Not IsKnownStatementDiv(parts(2)) And CInt(parts(4)) = targetYr And parts(5) = targetRepCode Then
                            AddUnique corpAccounts(compKey), parts(2) & " | " & parts(3)
                        End If
                    End If
                ElseIf UBound(parts) = 4 Then
                    If parts(0) = compKey Then
                        If Not IsKnownStatementDiv(parts(1)) And CInt(parts(3)) = targetYr And parts(4) = targetRepCode Then
                            AddUnique corpAccounts(compKey), parts(1) & " | " & parts(2)
                        End If
                    End If
                End If
            Next cKey
        Next pRev

    Next compKey

    ' 10. Write results to Excel using 2D Variant Arrays (Fast batch writing)
    Dim startCol As Long
    startCol = 1

    For Each compKey In activeComps.Keys
        compInfo = activeComps(compKey)

        Dim sCode As String, cName As String, fDiv As String
        sCode = compInfo(0)
        cName = compInfo(1)
        fDiv = compInfo(2)

        ' Write Company Header Info in batch
        Dim metaArr(1 To 3, 1 To 2) As Variant
        metaArr(1, 1) = KoreanCompanyName(): metaArr(1, 2) = cName
        metaArr(2, 1) = KoreanStockCode(): metaArr(2, 2) = "A" & sCode
        metaArr(3, 1) = KoreanCategory()
        metaArr(3, 2) = IIf(fDiv = "CFS", KoreanCfsStatement(), KoreanOfsStatement())
        wsOut.Cells(1, startCol).Resize(3, 2).Value = metaArr

        ' Write Table Headers in batch
        Dim headerArr() As Variant
        ReDim headerArr(1 To 1, 1 To numPeriods + 2)
        headerArr(1, 1) = KoreanCategory()
        headerArr(1, 2) = KoreanAccountName()
        For p = 1 To numPeriods
            headerArr(1, p + 2) = periodLabels(p)
        Next p
        wsOut.Cells(5, startCol).Resize(1, numPeriods + 2).Value = headerArr

        ' Write Accounts and values in batch
        Dim accCol As Collection
        Set accCol = corpAccounts(compKey)
        Dim numRows As Long
        numRows = accCol.Count

        If numRows > 0 Then
            Dim blockData() As Variant
            ReDim blockData(1 To numRows, 1 To numPeriods + 2)

            Dim rIdx As Long
            rIdx = 1

            Dim accItem As Variant
            For Each accItem In accCol
                Dim accParts() As String
                accParts = Split(accItem, " | ")

                Dim sjDiv As String, accName As String
                sjDiv = accParts(0)
                accName = accParts(1)

                ' Labels
                blockData(rIdx, 1) = sjDiv
                blockData(rIdx, 2) = accName

                ' Values from cache (Fallback to old format)
                For p = 1 To numPeriods
                    Dim dataKeyNew As String, dataKeyOld As String
                    dataKeyNew = compKey & "|" & fDiv & "|" & sjDiv & "|" & accName & "|" & CStr(periodYears(p)) & "|" & periodRepCodes(p)
                    dataKeyOld = compKey & "|" & sjDiv & "|" & accName & "|" & CStr(periodYears(p)) & "|" & periodRepCodes(p)

                    If cacheDict.Exists(dataKeyNew) Then
                        blockData(rIdx, p + 2) = CleanAmount(cacheDict(dataKeyNew))
                    ElseIf cacheDict.Exists(dataKeyOld) Then
                        blockData(rIdx, p + 2) = CleanAmount(cacheDict(dataKeyOld))
                    Else
                        blockData(rIdx, p + 2) = ""
                    End If
                Next p

                rIdx = rIdx + 1
            Next accItem

            ' Write blockData to worksheet in one go
            wsOut.Cells(6, startCol).Resize(numRows, numPeriods + 2).Value = blockData
        End If

        ' Apply styling for this block
        FormatCompanyBlock wsOut, startCol, numPeriods, 5 + numRows

        ' Advance columns (numPeriods + 2 data columns, and 1 blank separator column)
        startCol = startCol + numPeriods + 3
    Next compKey

    ' Delete DART_Cache sheet if it exists to clean up workbook
    On Error Resume Next
    Dim wsCache As Worksheet
    Set wsCache = wb.Worksheets("DART_Cache")
    If Not wsCache Is Nothing Then
        wsCache.Delete
    End If
    On Error GoTo 0

    ' Final auto-fit
    wsOut.Columns.AutoFit

    ' Restore states
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "DART financial table download complete. Processed " & compIdx & " companies.", vbInformation
    Exit Sub

ParseError:
    MsgBox "Failed to parse target date in MAIN!C3. Ensure format is like '2026" & KoreanYear() & " " & KoreanQuarter1() & "'.", vbCritical
    Exit Sub

CleanFail:
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Download failed: " & Err.Description, vbCritical
End Sub

' Parses dates like "2026 year quarter 1"
Private Sub ParseTargetDate(ByVal dateStr As String, ByRef outYear As Integer, ByRef outReprtCode As String, ByRef outPeriodName As String)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "([0-9]{4})"

    Dim matches As Object
    Set matches = regEx.Execute(dateStr)
    If matches.Count > 0 Then
        outYear = CInt(matches(0).SubMatches(0))
    Else
        Err.Raise vbObjectError + 2001, , "Invalid year format in date string."
    End If

    If InStr(dateStr, KoreanQuarter1()) > 0 Or InStr(dateStr, "1Q") > 0 Then
        outReprtCode = "11013"
        outPeriodName = KoreanQuarter1()
    ElseIf InStr(dateStr, KoreanHalf()) > 0 Or InStr(dateStr, KoreanQuarter2()) > 0 Or InStr(dateStr, "2Q") > 0 Then
        outReprtCode = "11012"
        outPeriodName = KoreanHalf()
    ElseIf InStr(dateStr, KoreanQuarter3()) > 0 Or InStr(dateStr, "3Q") > 0 Then
        outReprtCode = "11014"
        outPeriodName = KoreanQuarter3()
    ElseIf InStr(dateStr, KoreanQuarter4()) > 0 Or InStr(dateStr, "4Q") > 0 Or InStr(dateStr, KoreanAnnual()) > 0 Or InStr(dateStr, KoreanBusiness()) > 0 Then
        outReprtCode = "11011"
        outPeriodName = KoreanAnnual()
    Else
        ' Fallback to annual
        outReprtCode = "11011"
        outPeriodName = KoreanAnnual()
    End If
End Sub

' Downloads and builds the KOSPI/KOSDAQ listed company corp code map in DART_CorpCodes sheet
' Downloads and builds the KOSPI/KOSDAQ listed company corp code map in the DART corp-code sheet
Public Sub InitializeCorpCodes()
    Dim wb As Workbook
    Dim wsMain As Worksheet
    Dim wsCorp As Worksheet
    Dim apiKey As String
    Dim fso As Object

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = ThisWorkbook

    LogMsg "InitializeCorpCodes: Started"

    On Error Resume Next
    Set wsMain = wb.Worksheets("MAIN")
    On Error GoTo 0

    If wsMain Is Nothing Then
        MsgBox "MAIN sheet was not found. Please create a sheet named 'MAIN'.", vbCritical
        Exit Sub
    End If

    apiKey = Trim$(CStr(wsMain.Range("C2").Value))
    If Len(apiKey) = 0 Then
        MsgBox "API Key in cell C2 of MAIN sheet is empty.", vbCritical
        Exit Sub
    End If

    ' Create or clear the DART corp-code worksheet
    On Error Resume Next
    Set wsCorp = wb.Worksheets(DartCorpCodeSheetName())
    On Error GoTo 0

    If wsCorp Is Nothing Then
        Set wsCorp = wb.Worksheets.Add(After:=wsMain)
        wsCorp.Name = DartCorpCodeSheetName()
    Else
        wsCorp.UsedRange.Clear
    End If

    MsgBox "First-time setup: Downloading Dart corporate codes... This may take a few seconds.", vbInformation

    ' Download zip
    Dim cachedXmlPath As String
    cachedXmlPath = Environ("TEMP") & "\CORPCODE.xml"

    Dim url As String
    url = BASE_API_URL & "corpCode.xml?crtfc_key=" & apiKey

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        MsgBox "Failed to download corpCode zip: HTTP " & http.Status, vbCritical
        Exit Sub
    End If

    Dim zipPath As String
    zipPath = Environ("TEMP") & "\corpCode.zip"

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile zipPath, 2 ' Overwrite
    stream.Close

    ' Unzip CORPCODE.xml
    Dim tempDir As String
    tempDir = Environ("TEMP") & "\corpCode_extracted"
    UnzipFile zipPath, tempDir

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(tempDir & "\CORPCODE.xml") Then
        If fso.FileExists(cachedXmlPath) Then fso.DeleteFile cachedXmlPath, True
        fso.CopyFile tempDir & "\CORPCODE.xml", cachedXmlPath
    Else
        MsgBox "CORPCODE.xml not found in ZIP.", vbCritical
        Exit Sub
    End If

    ' Clean up zip files
    On Error Resume Next
    fso.DeleteFile zipPath, True
    fso.DeleteFolder tempDir, True
    On Error GoTo 0

    LogMsg "InitializeCorpCodes: Reading CORPCODE.xml"
    Dim xmlContent As String
    xmlContent = ReadUTF8File(cachedXmlPath)
    LogMsg "InitializeCorpCodes: Loaded XML, length = " & Len(xmlContent)

    ' Split by <list> tag to process each company block
    Dim parts() As String
    parts = Split(xmlContent, "<list>")
    LogMsg "InitializeCorpCodes: Split into " & UBound(parts) & " blocks, starting parsing loop"

    ' Pre-allocate results array to size 10000 (enough for ~4000 listed companies)
    Dim results() As Variant
    ReDim results(1 To 10000, 1 To 3)
    Dim matchCount As Long
    matchCount = 0

    Dim i As Long
    Dim part As String
    Dim pStock As Long, pStockEnd As Long
    Dim pCorp As Long, pCorpEnd As Long
    Dim pName As Long, pNameEnd As Long
    Dim currentStockCode As String
    Dim currentCorpCode As String
    Dim currentCorpName As String

    For i = 1 To UBound(parts)
        part = parts(i)
        pStock = InStr(part, "<stock_code>")
        if pStock > 0 Then
            pStockEnd = InStr(pStock + 12, part, "</stock_code>")
            If pStockEnd > pStock + 12 Then
                currentStockCode = Trim$(Mid$(part, pStock + 12, pStockEnd - (pStock + 12)))
                ' Listed companies have exactly 6-digit stock codes
                If Len(currentStockCode) = 6 Then
                    pCorp = InStr(part, "<corp_code>")
                    If pCorp > 0 Then
                        pCorpEnd = InStr(pCorp + 11, part, "</corp_code>")
                        If pCorpEnd > pCorp + 11 Then
                            currentCorpCode = Trim$(Mid$(part, pCorp + 11, pCorpEnd - (pCorp + 11)))
                        End If
                    End If

                    pName = InStr(part, "<corp_name>")
                    If pName > 0 Then
                        pNameEnd = InStr(pName + 11, part, "</corp_name>")
                        If pNameEnd > pName + 11 Then
                            currentCorpName = Trim$(Mid$(part, pName + 11, pNameEnd - (pName + 11)))
                        End If
                    End If

                    matchCount = matchCount + 1
                    results(matchCount, 1) = "'" & currentStockCode
                    results(matchCount, 2) = "'" & currentCorpCode
                    results(matchCount, 3) = currentCorpName
                End If
            End If
        End If
    Next i
    LogMsg "InitializeCorpCodes: Parsing loop finished, matchCount = " & matchCount

    ' Write to sheet in batch
    wsCorp.Cells(1, 1).Value = "Stock Code"
    wsCorp.Cells(1, 2).Value = "Corp Code"
    wsCorp.Cells(1, 3).Value = "Corp Name"

    Dim finalResults() As Variant
    If matchCount > 0 Then
        ReDim finalResults(1 To matchCount, 1 To 3)
        Dim idx As Long
        For idx = 1 To matchCount
            finalResults(idx, 1) = results(idx, 1)
            finalResults(idx, 2) = results(idx, 2)
            finalResults(idx, 3) = results(idx, 3)
        Next idx
        wsCorp.Cells(2, 1).Resize(matchCount, 3).Value = finalResults
        LogMsg "InitializeCorpCodes: Wrote " & matchCount & " listed companies to DART corp-code sheet"
    End If

    ' Delete CORPCODE.xml to save disk space
    On Error Resume Next
    fso.DeleteFile cachedXmlPath, True
    On Error GoTo 0

    wsCorp.Columns.AutoFit
    LogMsg "InitializeCorpCodes: Finished"
    MsgBox DartCorpCodeInitDoneMessage(matchCount), vbInformation
End Sub

' Loads corporate code mapping from the DART corp-code sheet
Private Sub LoadCorpCodeMap(ByRef corpMap As Object, ByRef nameMap As Object)
    LogMsg "LoadCorpCodeMap: Started"
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim wsCorp As Worksheet
    On Error Resume Next
    Set wsCorp = wb.Worksheets(DartCorpCodeSheetName())
    On Error GoTo 0

    If wsCorp Is Nothing Then
        MsgBox DartCorpCodeMissingMessage(), vbCritical
        Err.Raise vbObjectError + 2004, , DartCorpCodeMissingDescription()
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsCorp.Cells(wsCorp.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox DartCorpCodeEmptyMessage(), vbCritical
        Err.Raise vbObjectError + 2005, , DartCorpCodeEmptyDescription()
        Exit Sub
    End If

    LogMsg "LoadCorpCodeMap: Loading cached mapping from DART corp-code sheet"
    ' Load mapping directly from the worksheet using 2D Variant Array (High Performance)
    Dim codeData As Variant
    codeData = wsCorp.Range(wsCorp.Cells(1, 1), wsCorp.Cells(lastRow, 3)).Value

    Dim i As Long
    Dim sCode As String, cCode As String, cName As String
    For i = 2 To UBound(codeData, 1)
        sCode = Trim$(CStr(codeData(i, 1)))
        cCode = Trim$(CStr(codeData(i, 2)))
        cName = Trim$(CStr(codeData(i, 3)))

        If Len(sCode) < 6 And Len(sCode) > 0 Then sCode = String(6 - Len(sCode), "0") & sCode
        If Len(cCode) < 8 And Len(cCode) > 0 Then cCode = String(8 - Len(cCode), "0") & cCode

        If Len(sCode) = 6 And Len(cCode) = 8 Then
            corpMap(sCode) = cCode
            nameMap(cCode) = cName
        End If
    Next i
    LogMsg "LoadCorpCodeMap: Loaded " & corpMap.Count & " cached company mappings from sheet"
    LogMsg "LoadCorpCodeMap: Exiting Sub"
End Sub

' Extracts tag value from single XML line
Private Function ExtractTagValue(ByVal line As String, ByVal tagName As String) As String
    Dim startTag As String, endTag As String
    startTag = "<" & tagName & ">"
    endTag = "</" & tagName & ">"

    Dim p1 As Long, p2 As Long
    p1 = InStr(line, startTag)
    If p1 > 0 Then
        p2 = InStr(p1 + Len(startTag), line, endTag)
        If p2 > 0 Then
            ExtractTagValue = Mid$(line, p1 + Len(startTag), p2 - (p1 + Len(startTag)))
            Exit Function
        End If
    End If
    ExtractTagValue = ""
End Function

' Unzips zipPath to destDir using Windows Shell API
Private Sub UnzipFile(ByVal zipPath As String, ByVal destDir As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(destDir) Then
        fso.DeleteFolder destDir, True
    End If
    fso.CreateFolder destDir

    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")

    Dim zipFolder As Object
    Set zipFolder = shellApp.Namespace(CVar(zipPath))
    Dim destFolder As Object
    Set destFolder = shellApp.Namespace(CVar(destDir))

    If zipFolder Is Nothing Or destFolder Is Nothing Then
        Err.Raise vbObjectError + 4001, , "Failed to initialize Shell Namespace."
    End If

    destFolder.CopyHere zipFolder.Items, 20

    Dim xmlPath As String
    xmlPath = destDir & "\CORPCODE.xml"

    Dim i As Integer
    For i = 1 To 100
        If fso.FileExists(xmlPath) Then Exit Sub
        Sleep 100
        DoEvents
    Next i

    If Not fso.FileExists(xmlPath) Then
        Err.Raise vbObjectError + 4002, , "Failed to unzip CORPCODE.xml (Timeout)."
    End If
End Sub

' Loads financial statement cache from DART_Cache sheet

' Batch writes newly added cache keys to DART_Cache sheet

' Fetches period financial statement in batch using multi-company API or single-company API if not present in cache
Private Sub FetchPeriodDataBatch(ByVal apiKey As String, ByVal corpCodes As Collection, ByVal yr As Integer, ByVal reprtCode As String, ByRef cacheDict As Object, ByVal isMajor As Boolean, ByRef activeComps As Object)
    If corpCodes.Count = 0 Then Exit Sub

    If isMajor Then
        Dim chunkCodes As Collection
        Set chunkCodes = New Collection

        Dim i As Long
        For i = 1 To corpCodes.Count
            chunkCodes.Add corpCodes(i)

            ' Chunk of 100 or last item
            If chunkCodes.Count = 100 Or i = corpCodes.Count Then
                SendBatchRequest apiKey, chunkCodes, yr, reprtCode, cacheDict, activeComps
                Set chunkCodes = New Collection
            End If
        Next i
    Else
        Dim j As Long
        Dim compKey As String
        Dim compInfo As Variant
        Dim fsDiv As String
        For j = 1 To corpCodes.Count
            compKey = corpCodes(j)
            compInfo = activeComps(compKey)
            fsDiv = compInfo(2)

            On Error Resume Next
            Dim fetchStatus As String
            If fsDiv = "AUTO" Then
                fetchStatus = SendSingleRequest(apiKey, compKey, "CFS", yr, reprtCode, cacheDict, activeComps)
                If Err.Number = 0 And fetchStatus = "013" Then
                    LogMsg "Auto fs fallback: corp=" & compKey & " year=" & yr & " reprt=" & reprtCode & " CFS returned 013; trying OFS"
                    fetchStatus = SendSingleRequest(apiKey, compKey, "OFS", yr, reprtCode, cacheDict, activeComps)
                End If
            Else
                fetchStatus = SendSingleRequest(apiKey, compKey, fsDiv, yr, reprtCode, cacheDict, activeComps)
            End If
            If Err.Number <> 0 Then
                LogMsg "Single fetch skipped: corp=" & compKey & " fsDiv=" & fsDiv & " year=" & yr & " reprt=" & reprtCode & " error=" & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
            Sleep 200 ' 0.2s delay to prevent temporary API blocking
        Next j
    End If
End Sub

' Sends the actual HTTP request and caches the results (For Major Accounts)
Private Sub SendBatchRequest(ByVal apiKey As String, ByVal chunkCodes As Collection, ByVal yr As Integer, ByVal reprtCode As String, ByRef cacheDict As Object, ByRef activeComps As Object)
    Dim corpCodesStr As String
    corpCodesStr = ""

    Dim item As Variant
    For Each item In chunkCodes
        If Len(corpCodesStr) > 0 Then
            corpCodesStr = corpCodesStr & "," & CStr(item)
        Else
            corpCodesStr = CStr(item)
        End If
    Next item

    Dim url As String
    url = BASE_API_URL & "fnlttMultiAcnt.xml?crtfc_key=" & apiKey & "&corp_code=" & corpCodesStr & "&bsns_year=" & yr & "&reprt_code=" & reprtCode

    LogMsg "SendBatchRequest: Fetching URL: " & url
    Dim status As String
    Dim xmlDoc As Object
    On Error GoTo FetchError
    Set xmlDoc = FetchDartXml(url, status)
    On Error GoTo 0
    LogMsg "SendBatchRequest: Fetch status = " & status

    Dim cCode As Variant
    If status = "000" Then
        Dim listNodes As Object
        Set listNodes = xmlDoc.SelectNodes("/result/list")

        Dim node As Object
        Dim xmlCorpCode As String, xmlFsDiv As String, sjDiv As String, accName As String
        Dim thAmt As String, frAmt As String, bfeAmt As String

        For Each node In listNodes
            xmlCorpCode = Trim$(GetNodeTextSafe(node, "corp_code"))
            xmlFsDiv = Trim$(GetNodeTextSafe(node, "fs_div"))
            sjDiv = NormalizeStatementDiv(GetNodeTextSafe(node, "sj_div"))
            accName = Trim$(GetNodeTextSafe(node, "account_nm"))

            If Len(xmlCorpCode) > 0 Then
                If reprtCode = "11011" Then ' Annual
                    thAmt = GetNodeTextSafe(node, "thstrm_amount")
                    frAmt = GetNodeTextSafe(node, "frmtrm_amount")
                    bfeAmt = GetNodeTextSafe(node, "bfefrmtrm_amount")

                    SaveCacheValue cacheDict, xmlCorpCode, xmlFsDiv, sjDiv, accName, yr, "11011", thAmt
                    SaveCacheValue cacheDict, xmlCorpCode, xmlFsDiv, sjDiv, accName, yr - 1, "11011", frAmt
                    SaveCacheValue cacheDict, xmlCorpCode, xmlFsDiv, sjDiv, accName, yr - 2, "11011", bfeAmt
                Else ' Quarter
                    thAmt = GetNodeTextSafe(node, "thstrm_amount")

                    SaveCacheValue cacheDict, xmlCorpCode, xmlFsDiv, sjDiv, accName, yr, reprtCode, thAmt
                End If
            End If
        Next node

        ' Mark all requested companies in this chunk as successfully fetched for both CFS and OFS
        For Each cCode In chunkCodes
            MarkFetched cacheDict, CStr(cCode), "CFS", yr, reprtCode
            MarkFetched cacheDict, CStr(cCode), "OFS", yr, reprtCode
        Next cCode

    ElseIf status = "010" Or status = "011" Then
        ' Authentication error: Do not cache status, throw error
        Dim compDetails As String
        compDetails = ""
        Dim compName As String, stockCode As String
        For Each cCode In chunkCodes
            compName = "Unknown"
            stockCode = "Unknown"
            If Not activeComps Is Nothing Then
                If activeComps.Exists(CStr(cCode)) Then
                    stockCode = activeComps(CStr(cCode))(0)
                    compName = activeComps(CStr(cCode))(1)
                End If
            End If
            If Len(compDetails) > 0 Then
                compDetails = compDetails & ", "
            End If
            compDetails = compDetails & compName & "(" & stockCode & ")"
        Next cCode

        Dim statusDesc As String
        Select Case status
            Case "010": statusDesc = KoreanUnregisteredKey()
            Case "011": statusDesc = KoreanInvalidKey()
            Case Else: statusDesc = KoreanOtherError()
        End Select

        Err.Raise vbObjectError + 3010, , "DART Key Error: " & status & " (" & statusDesc & ")" & vbCrLf & _
                                         KoreanTargetCompany() & compDetails
    Else
        ' "013"/"020" No Data or other status codes: Cache as fetched so we do not query again
        For Each cCode In chunkCodes
            MarkFetched cacheDict, CStr(cCode), "CFS", yr, reprtCode
            MarkFetched cacheDict, CStr(cCode), "OFS", yr, reprtCode
        Next cCode
    End If

    Exit Sub

FetchError:
    LogMsg "SendBatchRequest FetchError: batch " & corpCodesStr & " / " & yr & " / " & reprtCode & " Error: " & Err.Description
    Dim fbCode As Variant

    If chunkCodes.Count > 1 Then
        Dim singleChunk As Collection
        For Each fbCode In chunkCodes
            Set singleChunk = New Collection
            singleChunk.Add fbCode
            SendBatchRequest apiKey, singleChunk, yr, reprtCode, cacheDict, activeComps
        Next fbCode
    Else
        ' Only the failed company/period is treated as blank.
        On Error Resume Next
        For Each fbCode In chunkCodes
            MarkFetched cacheDict, CStr(fbCode), "CFS", yr, reprtCode
            MarkFetched cacheDict, CStr(fbCode), "OFS", yr, reprtCode
        Next fbCode
        On Error GoTo 0
    End If
End Sub

' Sends the actual HTTP request and caches the results (For All Detailed Accounts)
Private Function SendSingleRequest(ByVal apiKey As String, ByVal corpCode As String, ByVal fsDiv As String, ByVal yr As Integer, ByVal reprtCode As String, ByRef cacheDict As Object, ByRef activeComps As Object) As String
    Dim url As String
    url = BASE_API_URL & "fnlttSinglAcntAll.xml?crtfc_key=" & apiKey & "&corp_code=" & corpCode & "&bsns_year=" & yr & "&reprt_code=" & reprtCode & "&fs_div=" & fsDiv

    LogMsg "SendSingleRequest: Fetching URL: " & url
    Dim status As String
    Dim xmlDoc As Object
    On Error GoTo FetchError
    Set xmlDoc = FetchDartXml(url, status)
    On Error GoTo 0
    LogMsg "SendSingleRequest: Fetch status = " & status

    If status = "000" Then
        Dim listNodes As Object
        Set listNodes = xmlDoc.SelectNodes("/result/list")

        Dim node As Object
        Dim xmlCorpCode As String, xmlFsDiv As String, sjDiv As String, accName As String
        Dim thAmt As String, frAmt As String, bfeAmt As String

        For Each node In listNodes
            xmlCorpCode = Trim$(GetNodeTextSafe(node, "corp_code"))
            xmlFsDiv = Trim$(GetNodeTextSafe(node, "fs_div"))
            If Len(xmlFsDiv) = 0 Then xmlFsDiv = fsDiv
            sjDiv = NormalizeStatementDiv(GetNodeTextSafe(node, "sj_div"))
            accName = Trim$(GetNodeTextSafe(node, "account_nm"))

            If Len(xmlCorpCode) > 0 And Len(sjDiv) > 0 And Len(accName) > 0 Then
                If reprtCode = "11011" Then ' Annual
                    thAmt = GetNodeTextSafe(node, "thstrm_amount")
                    frAmt = GetNodeTextSafe(node, "frmtrm_amount")
                    bfeAmt = GetNodeTextSafe(node, "bfefrmtrm_amount")

                    SaveCacheValue cacheDict, xmlCorpCode, xmlFsDiv, sjDiv, accName, yr, "11011", thAmt
                    SaveCacheValue cacheDict, xmlCorpCode, xmlFsDiv, sjDiv, accName, yr - 1, "11011", frAmt
                    SaveCacheValue cacheDict, xmlCorpCode, xmlFsDiv, sjDiv, accName, yr - 2, "11011", bfeAmt
                Else ' Quarter
                    thAmt = GetNodeTextSafe(node, "thstrm_amount")

                    SaveCacheValue cacheDict, xmlCorpCode, xmlFsDiv, sjDiv, accName, yr, reprtCode, thAmt
                End If
            End If
        Next node

        ' Mark as successfully fetched
        MarkFetched cacheDict, corpCode, fsDiv, yr, reprtCode
        SendSingleRequest = status

    ElseIf status = "010" Or status = "011" Then
        ' Authentication error: Do not cache status, throw error
        Dim singleCompName As String, singleStockCode As String
        singleCompName = "Unknown"
        singleStockCode = "Unknown"
        If Not activeComps Is Nothing Then
            If activeComps.Exists(corpCode) Then
                singleStockCode = activeComps(corpCode)(0)
                singleCompName = activeComps(corpCode)(1)
            End If
        End If

        Dim singleStatusDesc As String
        Select Case status
            Case "010": singleStatusDesc = KoreanUnregisteredKey()
            Case "011": singleStatusDesc = KoreanInvalidKey()
            Case Else: singleStatusDesc = KoreanOtherError()
        End Select

        Err.Raise vbObjectError + 3010, , "DART Key Error: " & status & " (" & singleStatusDesc & ")" & vbCrLf & _
                                         KoreanTargetCompany() & singleCompName & "(" & singleStockCode & ")"
    Else
        ' "013"/"020" No Data or other status codes: Cache as fetched so we do not query again
        MarkFetched cacheDict, corpCode, fsDiv, yr, reprtCode
        SendSingleRequest = status
    End If

    Exit Function

FetchError:
    LogMsg "SendSingleRequest FetchError: corp " & corpCode & " / " & yr & " / " & reprtCode & " Error: " & Err.Description
    ' Fallback: mark company as fetched (empty) on error
    On Error Resume Next
    MarkFetched cacheDict, corpCode, fsDiv, yr, reprtCode
    On Error GoTo 0
    SendSingleRequest = "ERROR"
End Function

' Normalizes DART statement division codes for display and cache keys
Private Function NormalizeStatementDiv(ByVal sjDiv As String) As String
    NormalizeStatementDiv = UCase$(Trim$(sjDiv))
End Function

' Returns the desired output statement order: BS, IS, CIS, CF, SCE
Private Function StatementDivByOrder(ByVal orderIdx As Integer) As String
    Select Case orderIdx
        Case 1
            StatementDivByOrder = "BS"
        Case 2
            StatementDivByOrder = "IS"
        Case 3
            StatementDivByOrder = "CIS"
        Case 4
            StatementDivByOrder = "CF"
        Case 5
            StatementDivByOrder = "SCE"
        Case Else
            StatementDivByOrder = ""
    End Select
End Function

' True for the statement divisions we place in a fixed order.
Private Function IsKnownStatementDiv(ByVal sjDiv As String) As Boolean
    Select Case UCase$(Trim$(sjDiv))
        Case "BS", "IS", "CIS", "CF", "SCE"
            IsKnownStatementDiv = True
    End Select
End Function

' AUTO mode accepts whichever statement type was successfully fetched.
Private Function FsDivMatches(ByVal cachedFsDiv As String, ByVal requestedFsDiv As String) As Boolean
    If requestedFsDiv = "AUTO" Then
        FsDivMatches = (cachedFsDiv = "CFS" Or cachedFsDiv = "OFS")
    Else
        FsDivMatches = (cachedFsDiv = requestedFsDiv)
    End If
End Function

Private Function ResolveAutoFsDiv(ByRef cacheDict As Object, ByVal corpCode As String) As String
    Dim key As Variant
    Dim parts() As String
    Dim hasCfs As Boolean
    Dim hasOfs As Boolean

    For Each key In cacheDict.Keys
        parts = Split(CStr(key), "|")
        If UBound(parts) = 5 And parts(0) = corpCode Then
            If parts(1) = "CFS" Then
                hasCfs = True
            ElseIf parts(1) = "OFS" Then
                hasOfs = True
            End If
        End If
    Next key

    If hasCfs Then
        ResolveAutoFsDiv = "CFS"
    ElseIf hasOfs Then
        ResolveAutoFsDiv = "OFS"
    Else
        ResolveAutoFsDiv = "CFS"
    End If
End Function

' Helper to write a value into cache dictionaries
Private Sub SaveCacheValue(ByRef cacheDict As Object, ByVal corpCode As String, ByVal fsDiv As String, ByVal sjDiv As String, ByVal accName As String, ByVal yr As Integer, ByVal repCode As String, ByVal amt As String)
    Dim key As String
    key = corpCode & "|" & fsDiv & "|" & sjDiv & "|" & accName & "|" & CStr(yr) & "|" & repCode
    cacheDict(key) = amt
' newCache(key) = amt (removed)
End Sub

' Helper to mark a query as fetched in cache dictionaries
Private Sub MarkFetched(ByRef cacheDict As Object, ByVal corpCode As String, ByVal fsDiv As String, ByVal yr As Integer, ByVal repCode As String)
    Dim key As String
    key = corpCode & "|" & fsDiv & "|" & CStr(yr) & "|" & repCode
    cacheDict(key) = "Fetched"
' newCache(key) = "Fetched" (removed)
End Sub

' Fetches URL and parses XML, returns status code
Private Function FetchDartXml(ByVal url As String, ByRef outStatus As String) As Object
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 3001, , "HTTP Error " & http.Status & " fetching OpenDART."
    End If

    ' Decode responseBody using UTF-8 to prevent Korean character corruption
    Dim responseText As String
    responseText = BytesToUTF8String(http.ResponseBody)

    ' Strip XML declaration to prevent MSXML loadXML UTF-8 encoding mismatch error
    If InStr(1, responseText, "<?xml", vbTextCompare) > 0 Then
        Dim pDeclarationEnd As Long
        pDeclarationEnd = InStr(1, responseText, "?>")
        If pDeclarationEnd > 0 Then
            responseText = Mid$(responseText, pDeclarationEnd + 2)
        End If
    End If

    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.Async = False
    xmlDoc.validateOnParse = False
    xmlDoc.resolveExternals = False

    ' Load the UTF-8 text into the XML document
    If Not xmlDoc.LoadXML(responseText) Then
        Err.Raise vbObjectError + 3002, , "Failed to load XML content: " & xmlDoc.parseError.reason
    End If

    Dim statusNode As Object
    Set statusNode = xmlDoc.SelectSingleNode("/result/status")
    If Not statusNode Is Nothing Then
        outStatus = statusNode.Text
    Else
        outStatus = "UNKNOWN"
    End If

    Set FetchDartXml = xmlDoc
End Function

' Safely reads text from an XML node
Private Function GetNodeTextSafe(ByVal parentNode As Object, ByVal nodeName As String) As String
    Dim childNode As Object
    Set childNode = parentNode.SelectSingleNode(nodeName)
    If Not childNode Is Nothing Then
        GetNodeTextSafe = Trim$(childNode.Text)
    Else
        GetNodeTextSafe = ""
    End If
End Function

' Normalizes stock codes (removes leading A if present and pads to 6 digits)
Private Function NormalizeStockCode(ByVal code As String) As String
    code = Trim$(code)
    If Left$(code, 1) = "A" Or Left$(code, 1) = "a" Then
        code = Mid$(code, 2)
    End If
    Do While Len(code) < 6
        code = "0" & code
    Loop
    NormalizeStockCode = code
End Function

' Cleans number formats to allow writing as double to Excel cells
Private Function CleanAmount(ByVal amtStr As String) As Variant
    amtStr = Replace(amtStr, ",", "")
    amtStr = Trim$(amtStr)

    If Len(amtStr) = 0 Or amtStr = "-" Then
        CleanAmount = ""
    Else
        If IsNumeric(amtStr) Then
            CleanAmount = CDbl(amtStr)
        Else
            CleanAmount = amtStr
        End If
    End If
End Function

' Adds unique string items to a VBA collection
Private Sub AddUnique(ByVal col As Collection, ByVal val As String)
    On Error Resume Next
    col.Add val, val
    On Error GoTo 0
End Sub

' Formats a single company table block inside the sheet
Private Sub FormatCompanyBlock(ByVal ws As Worksheet, ByVal startCol As Long, ByVal numP As Integer, ByVal endRow As Long)
    Dim cellRange As Range
    Dim lastCol As Long
    lastCol = startCol + 1 + numP

    ' Title headers
    ws.Range(ws.Cells(1, startCol), ws.Cells(3, lastCol)).Font.Name = "Malgun Gothic"
    ws.Range(ws.Cells(1, startCol), ws.Cells(3, startCol)).Font.Bold = True
    ws.Range(ws.Cells(1, startCol), ws.Cells(3, startCol)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, startCol), ws.Cells(3, startCol)).Interior.Color = RGB(242, 244, 247)

    ' Border for meta info block
    Set cellRange = ws.Range(ws.Cells(1, startCol), ws.Cells(3, lastCol))
    cellRange.Borders.LineStyle = xlContinuous
    cellRange.Borders.Color = RGB(218, 222, 229)

    ' Table Headers formatting
    Set cellRange = ws.Range(ws.Cells(5, startCol), ws.Cells(5, lastCol))
    cellRange.Font.Name = "Malgun Gothic"
    cellRange.Font.Bold = True
    cellRange.HorizontalAlignment = xlCenter
    cellRange.Interior.Color = RGB(74, 119, 202) ' Professional blue
    cellRange.Font.Color = RGB(255, 255, 255)

    ' Table Data cells formatting
    Set cellRange = ws.Range(ws.Cells(6, startCol), ws.Cells(endRow, lastCol))
    cellRange.Font.Name = "Malgun Gothic"
    cellRange.Font.Size = 10

    ' Align accounts and classification
    ws.Range(ws.Cells(6, startCol), ws.Cells(endRow, startCol)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(6, startCol + 1), ws.Cells(endRow, startCol + 1)).HorizontalAlignment = xlLeft

    Dim c As Long
    For c = startCol + 2 To lastCol
        Dim dataRange As Range
        Set dataRange = ws.Range(ws.Cells(6, c), ws.Cells(endRow, c))
        dataRange.HorizontalAlignment = xlRight
        dataRange.NumberFormat = "#,##0;(#,##0);""-"""
    Next c

    ' Grid borders
    Set cellRange = ws.Range(ws.Cells(5, startCol), ws.Cells(endRow, lastCol))
    cellRange.Borders.LineStyle = xlContinuous
    cellRange.Borders.Color = RGB(218, 222, 229)
End Sub

' Custom MsgBox wrapper to prevent hanging when Excel runs invisibly
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    If Application.Visible Then
        MsgBox = VBA.Interaction.MsgBox(Prompt, Buttons, Title)
    Else
        Debug.Print "MsgBox [" & Title & "]: " & Prompt
        MsgBox = vbOK
    End If
End Function

' Logging helper to write to a text file
Private Sub LogMsg(ByVal msg As String)
    On Error Resume Next
    Dim fNum As Integer
    fNum = FreeFile
    Dim logPath As String
    logPath = Environ("TEMP") & "\vba_log.txt"
    Open logPath For Append As #fNum
    Print #fNum, Now & " - " & msg
    Close #fNum
    On Error GoTo 0
End Sub

' Helper to read a UTF-8 file into a String using ADODB.Stream
Private Function ReadUTF8File(ByVal filePath As String) As String
    Dim stream As Object
    On Error GoTo CleanUp
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUTF8File = stream.ReadText(-1) ' adReadAll

CleanUp:
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
End Function

' Helper to decode UTF-8 byte arrays to a VBA string using ADODB.Stream
Private Function BytesToUTF8String(ByVal bytes As Variant) As String
    Dim stream As Object
    On Error GoTo CleanUp
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    BytesToUTF8String = stream.ReadText(-1)
CleanUp:
    If Err.Number <> 0 Then
        LogMsg "BytesToUTF8String Error: " & Err.Description
    End If
    If Not stream Is Nothing Then
        On Error Resume Next
        stream.Close
        Set stream = Nothing
        On Error GoTo 0
    End If
End Function
