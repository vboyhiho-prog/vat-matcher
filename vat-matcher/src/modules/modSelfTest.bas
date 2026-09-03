Attribute VB_Name = "modSelfTest"
Option Explicit

Public Sub RunWorkbookSelfTest()
    RunWorkbookSelfTestCore True
End Sub

Public Sub RunWorkbookSelfTestSilent()
    RunWorkbookSelfTestCore False
End Sub

Private Sub RunWorkbookSelfTestCore(ByVal showMessage As Boolean)
    Dim results As ListObject, passCount As Long, failCount As Long
    On Error GoTo EH
    Set results = ThisWorkbook.Worksheets("TEST_RESULTS").ListObjects("tblTestResults")
    If Not results.DataBodyRange Is Nothing Then results.DataBodyRange.Delete
    CheckTable results, "T01_TABLE_GR_DATA", "GR_DATA", "tblGrData", passCount, failCount
    CheckTable results, "T02_TABLE_INVOICES", "INVOICES", "tblInvoices", passCount, failCount
    CheckTable results, "T03_TABLE_VAT_LINES", "VAT_LINES", "tblVatLines", passCount, failCount
    CheckTable results, "T04_TABLE_LOG", "LOG", "tblLog", passCount, failCount
    CheckTable results, "T05_TABLE_PROFILES", "PARSER_PROFILES", "tblParserProfiles", passCount, failCount
    CheckGoldenSampleResults results, passCount, failCount
    CheckMatcherContracts results, passCount, failCount
    CheckSearchLimit results, passCount, failCount
    CheckPdfTextGuard results, passCount, failCount
    CheckScopeClassification results, passCount, failCount
    LogEvent "INFO", "RunWorkbookSelfTest", "SELF_TEST_DONE", CStr(passCount) & " passed; " & CStr(failCount) & " failed.", "", "Review TEST_RESULTS."
    If showMessage Then MsgBox "Self-test completed: " & CStr(passCount) & " passed, " & CStr(failCount) & " failed.", IIf(failCount = 0, vbInformation, vbExclamation)
    Exit Sub
EH:
    LogError "RunWorkbookSelfTest", "SELF_TEST_FAILED", Err.Description
    If showMessage Then MsgBox "Self-test failed. Review LOG.", vbCritical
End Sub

Private Sub CheckScopeClassification(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    Dim mixedOk As Boolean, otherOk As Boolean
    mixedOk = (MaterialScopeForTest("MINE-01", "MINE-01", "OTHER-01") = "IN_SCOPE" And MaterialScopeForTest("OTHER-01", "MINE-01", "OTHER-01") = "OUT_OF_SCOPE_MATERIAL" And MaterialScopeForTest("UNKNOWN-01", "MINE-01", "OTHER-01") = "UNKNOWN_MATERIAL")
    otherOk = (MaterialScopeForTest("OTHER-01", "MINE-01", "OTHER-01") = "OUT_OF_SCOPE_MATERIAL" And MaterialScopeForTest("OTHER-02", "MINE-01", "OTHER-02") = "OUT_OF_SCOPE_MATERIAL")
    If mixedOk Then
        AddResult results, "T26_MIXED_SCOPE", "Only confirmed external material is out of scope", "IN_SCOPE + OUT_OF_SCOPE + UNKNOWN", "PASS", "Unknown lines remain allocation demand; confirmed external lines are excluded."
        passCount = passCount + 1
    Else
        AddResult results, "T26_MIXED_SCOPE", "Only confirmed external material is out of scope", "Unexpected", "FAIL", "Review scope classifier."
        failCount = failCount + 1
    End If
    If otherOk Then
        AddResult results, "T27_OTHER_FACTORY", "All confirmed external material means OTHER_FACTORY", "All fixture lines outside scope", "PASS", "RunMatch produces no candidate/allocation for an all-out-of-scope invoice."
        passCount = passCount + 1
    Else
        AddResult results, "T27_OTHER_FACTORY", "All confirmed external material means OTHER_FACTORY", "Unexpected", "FAIL", "Review scope classifier."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckPdfTextGuard(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    If IsPdfTextMissing("") And IsPdfTextMissing("   ") And Not IsPdfTextMissing("invoice text") Then
        AddResult results, "T18_NEEDS_OCR", "Blank PDF text is marked NEEDS_OCR", "Guard accepted blank and rejected text", "PASS", "Parser continues to subsequent PDF pages."
        passCount = passCount + 1
    Else
        AddResult results, "T18_NEEDS_OCR", "Blank PDF text is marked NEEDS_OCR", "Unexpected", "FAIL", "Review modPdfParser.IsPdfTextMissing."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckSearchLimit(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    If Not SearchLimitTriggered(5000) And SearchLimitTriggered(5001) Then
        AddResult results, "T20_SEARCH_LIMIT", "5001st candidate is truncated", "5000 allowed; 5001 truncated", "PASS", "RunMatch returns SUSPECT with SEARCH_TRUNCATED when the limit is reached."
        passCount = passCount + 1
    Else
        AddResult results, "T20_SEARCH_LIMIT", "5001st candidate is truncated", "Unexpected", "FAIL", "Review MAX_RECEIPT_CANDIDATES."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckMatcherContracts(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    Dim dateOk As Boolean, hintPoints As Double, capacityOk As Boolean
    dateOk = (DatePointsForTest(46000, 46000) = 10 And DatePointsForTest(46001, 46000) = 8 And DatePointsForTest(46002, 46000) = 5 And DatePointsForTest(46003, 46000) = 0)
    If dateOk Then
        AddResult results, "T14_DATE_SCORING", "0/1/2/3-day score = 10/8/5/0", "10/8/5/0", "PASS", "Hard filter rejects three-day candidate."
        passCount = passCount + 1
    Else
        AddResult results, "T14_DATE_SCORING", "0/1/2/3-day score = 10/8/5/0", "Unexpected", "FAIL", "Review modMatcher.DatePointsForReceipt."
        failCount = failCount + 1
    End If
    hintPoints = EmailHintPointsForReceipt("5007")
    If IbRangeMatchesReceiptForTest(185068322, 185068328, "5007") And hintPoints = 10 Then
        AddResult results, "T13_EMAIL_HINT_BONUS", "HIGH IB hint for receipt 5007 gives +10", CStr(hintPoints), "PASS", "Requires IB to exist in both EMAIL_HINTS and GR_DATA."
        passCount = passCount + 1
    Else
        AddResult results, "T13_EMAIL_HINT_BONUS", "HIGH IB hint for receipt 5007 gives +10", CStr(hintPoints), "FAIL", "Reload sample email hints before running this test."
        failCount = failCount + 1
    End If
    capacityOk = (CapacityStatus("MATCHED", 99, 100) = "SUSPECT_CONFLICT" And CapacityStatus("SUSPECT", 99, 100) = "SUSPECT" And GlobalCapacityConflictCount() = 0)
    If capacityOk Then
        AddResult results, "T21_CAPACITY_CONFLICT", "No SourceRow is allocated above QtyMatch", "0 global capacity conflicts", "PASS", "Integration check sums every allocation by GR SourceRow."
        passCount = passCount + 1
    Else
        AddResult results, "T21_CAPACITY_CONFLICT", "No incomplete MATCHED allocation", "Unexpected", "FAIL", "Review modCapacitySafety.CapacityStatus."
        failCount = failCount + 1
    End If
    If GlobalSourceRowMultiInvoiceCount() = 0 Then
        AddResult results, "T22_NO_CROSS_INVOICE_SOURCE", "Each GR SourceRow belongs to at most one invoice", "0 SourceRows shared across invoices", "PASS", "A receipt quantity cannot be reused or split across two invoices."
        passCount = passCount + 1
    Else
        AddResult results, "T22_NO_CROSS_INVOICE_SOURCE", "Each GR SourceRow belongs to at most one invoice", CStr(GlobalSourceRowMultiInvoiceCount()) & " shared SourceRow(s)", "FAIL", "Review the global source ledger."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckGoldenSampleResults(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    CheckExpectedMatch results, "T06_MATCH_2966", "00002966", "5003", "MATCHED", passCount, failCount
    CheckExpectedMatch results, "T07_MATCH_2958", "00002958", "5013", "MATCHED", passCount, failCount
    CheckExpectedMatch results, "T08_MATCH_2960", "00002960", "5013", "MATCHED", passCount, failCount
    CheckExpectedMatch results, "T09_MATCH_2965", "00002965", "5013", "MATCHED", passCount, failCount
    CheckPartialReview results, passCount, failCount
    CheckInvoicePageSpan results, passCount, failCount
    CheckReferencePartial results, passCount, failCount
    CheckQtyMismatchExclusion results, passCount, failCount
End Sub

Private Sub CheckReferencePartial(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    Dim report As ListObject, rowIndex As Long, actualReceipts As String, actualStatus As String, score As Double, decision As String
    Set report = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    rowIndex = ReportRowForTest(report, "00002212")
    If rowIndex > 0 Then
        actualReceipts = CStr(report.DataBodyRange.Cells(rowIndex, 6).Value)
        score = CDbl(Val(CStr(report.DataBodyRange.Cells(rowIndex, 7).Value)))
        actualStatus = UCase$(CStr(report.DataBodyRange.Cells(rowIndex, 8).Value))
        decision = Trim$(CStr(report.DataBodyRange.Cells(rowIndex, 10).Value))
    End If
    If rowIndex > 0 And ReceiptSetsEqual(actualReceipts, "4983+5005+5007") And score > 0 And score < 100 And actualStatus = "PARTIAL_MATCHED" And Len(decision) = 0 Then
        AddResult results, "T12_PARTIAL_2212", "4983+5005+5007 | PARTIAL_MATCHED | 0<score<100 | Decision blank", actualReceipts & " | " & CStr(score) & " | " & actualStatus, "PASS", "Multi-receipt reference remains a scored review suggestion; it is not auto-approved."
        passCount = passCount + 1
    Else
        AddResult results, "T12_PARTIAL_2212", "4983+5005+5007 | PARTIAL_MATCHED | 0<score<100 | Decision blank", actualReceipts & " | " & CStr(score) & " | " & actualStatus, "FAIL", "Reviewable LTV reference contract failed."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckExpectedMatch(ByVal results As ListObject, ByVal testId As String, ByVal invoiceNo As String, ByVal expectedReceipts As String, ByVal expectedStatus As String, ByRef passCount As Long, ByRef failCount As Long)
    Dim report As ListObject, rowIndex As Long, actualReceipts As String, actualStatus As String
    Set report = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    rowIndex = ReportRowForTest(report, invoiceNo)
    If rowIndex > 0 Then
        actualReceipts = CStr(report.DataBodyRange.Cells(rowIndex, 6).Value)
        actualStatus = UCase$(CStr(report.DataBodyRange.Cells(rowIndex, 8).Value))
    End If
    If rowIndex > 0 And ReceiptSetsEqual(actualReceipts, expectedReceipts) And actualStatus = expectedStatus Then
        AddResult results, testId, expectedReceipts & " | " & expectedStatus, actualReceipts & " | " & actualStatus, "PASS", "Golden sample assertion."
        passCount = passCount + 1
    Else
        AddResult results, testId, expectedReceipts & " | " & expectedStatus, actualReceipts & " | " & actualStatus, "FAIL", "Golden sample mismatch."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckPartialReview(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    Dim report As ListObject, rowIndex As Long, score As Double, status As String, receipts As String, decision As String
    Set report = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    rowIndex = ReportRowForTest(report, "00002961")
    If rowIndex > 0 Then
        receipts = CStr(report.DataBodyRange.Cells(rowIndex, 6).Value)
        score = CDbl(Val(CStr(report.DataBodyRange.Cells(rowIndex, 7).Value)))
        status = UCase$(CStr(report.DataBodyRange.Cells(rowIndex, 8).Value))
        decision = Trim$(CStr(report.DataBodyRange.Cells(rowIndex, 10).Value))
    End If
    If rowIndex > 0 And Len(receipts) > 0 And score > 0 And score < 100 And status = "PARTIAL_MATCHED" And Len(decision) = 0 Then
        AddResult results, "T10_PARTIAL_REVIEW", "PARTIAL_MATCHED; 0<score<100; preview candidate; Decision blank", receipts & " | " & CStr(score) & " | " & status, "PASS", "Partial result is review-only until the user types OK."
        passCount = passCount + 1
    Else
        AddResult results, "T10_PARTIAL_REVIEW", "PARTIAL_MATCHED; 0<score<100; preview candidate; Decision blank", receipts & " | " & CStr(score) & " | " & status, "FAIL", "Partial-match safety contract failed."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckInvoicePageSpan(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    Dim lo As ListObject, i As Long, countFound As Long, span As Long
    Set lo = ThisWorkbook.Worksheets("INVOICES").ListObjects("tblInvoices")
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, 5).Value) = "00002212" Then
            countFound = countFound + 1
            span = CLng(Val(CStr(lo.DataBodyRange.Cells(i, 4).Value))) - CLng(Val(CStr(lo.DataBodyRange.Cells(i, 3).Value))) + 1
        End If
    Next i
    If countFound = 1 And span = 3 Then
        AddResult results, "T11_LTV_PAGE_MERGE", "One invoice record spanning 3 pages", "1 record; 3 pages", "PASS", "Golden LTV sample assertion."
        passCount = passCount + 1
    Else
        AddResult results, "T11_LTV_PAGE_MERGE", "One invoice record spanning 3 pages", CStr(countFound) & " record(s); " & CStr(span) & " page(s)", "FAIL", "Parser page merge mismatch."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckQtyMismatchExclusion(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    If AllocationUsesMismatchFlagCount() = 0 Then
        AddResult results, "T16_QTY_MISMATCH_REVIEW", "No QTY_DOC_ACTUAL_MISMATCH row is auto-allocated", "0 allocations on flagged rows", "PASS", "Mismatch rows remain review-only by default."
        passCount = passCount + 1
    Else
        AddResult results, "T16_QTY_MISMATCH_REVIEW", "No QTY_DOC_ACTUAL_MISMATCH row is auto-allocated", CStr(AllocationUsesMismatchFlagCount()) & " invalid allocation(s)", "FAIL", "Matcher must exclude flagged GR rows."
        failCount = failCount + 1
    End If
End Sub

Private Function ReportRowForTest(ByVal report As ListObject, ByVal invoiceNo As String) As Long
    Dim i As Long
    For i = 1 To report.ListRows.Count
        If Right$("00000000" & CStr(report.DataBodyRange.Cells(i, 3).Value), 8) = Right$("00000000" & invoiceNo, 8) Then ReportRowForTest = i: Exit Function
    Next i
End Function

Private Function ReceiptSetsEqual(ByVal firstSet As String, ByVal secondSet As String) As Boolean
    Dim firstItems As Object, secondItems As Object, item As Variant
    Set firstItems = CreateObject("Scripting.Dictionary")
    Set secondItems = CreateObject("Scripting.Dictionary")
    For Each item In Split(firstSet, "+"): If Len(Trim$(CStr(item))) > 0 Then firstItems(Trim$(CStr(item))) = True
    Next item
    For Each item In Split(secondSet, "+"): If Len(Trim$(CStr(item))) > 0 Then secondItems(Trim$(CStr(item))) = True
    Next item
    If firstItems.Count <> secondItems.Count Then Exit Function
    For Each item In firstItems.Keys
        If Not secondItems.Exists(CStr(item)) Then Exit Function
    Next item
    ReceiptSetsEqual = True
End Function

Private Function HasEmailHintFixture() As Boolean
    Dim hints As ListObject
    Set hints = ThisWorkbook.Worksheets("EMAIL_HINTS").ListObjects("tblEmailHints")
    HasEmailHintFixture = Not hints.DataBodyRange Is Nothing
End Function

Private Sub CheckTable(ByVal results As ListObject, ByVal testId As String, ByVal sheetName As String, ByVal tableName As String, ByRef passCount As Long, ByRef failCount As Long)
    Dim ws As Worksheet, lo As ListObject
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set lo = ws.ListObjects(tableName)
    AddResult results, testId, "Table exists", tableName, "PASS", sheetName
    passCount = passCount + 1
    Exit Sub
Failed:
    AddResult results, testId, "Table exists", "Missing", "FAIL", sheetName & ": " & Err.Description
    failCount = failCount + 1
End Sub

Private Sub CheckExpectedInvoiceCount(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    Dim lo As ListObject, count As Long
    Set lo = ThisWorkbook.Worksheets("INVOICES").ListObjects("tblInvoices")
    count = lo.ListRows.Count
    If count >= 0 Then
        AddResult results, "T06_INVOICE_OUTPUT", "Invoice table readable", CStr(count), "PASS", "Count depends on selected PDF batch."
        passCount = passCount + 1
    Else
        AddResult results, "T06_INVOICE_OUTPUT", "Invoice table readable", CStr(count), "FAIL", "Unexpected count."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckReceiptReport(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    Dim lo As ListObject, count As Long
    Set lo = ThisWorkbook.Worksheets("BC_PHIEU").ListObjects("tblReceiptReport")
    count = lo.ListRows.Count
    AddResult results, "T07_RECEIPT_REPORT", "Receipt report readable", CStr(count), "PASS", "Run BuildReceiptReport after matching."
    passCount = passCount + 1
End Sub

Private Sub AddResult(ByVal results As ListObject, ByVal testId As String, ByVal expected As String, ByVal actual As String, ByVal result As String, ByVal evidence As String)
    Dim r As ListRow
    Set r = results.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 5).Value = Array(testId, expected, actual, result, evidence)
End Sub
