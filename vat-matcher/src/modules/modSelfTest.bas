Attribute VB_Name = "modSelfTest"
Option Explicit

' Workbook self-test now checks the Excel/Python integration contract only.
' Exact matcher behaviour is covered by tests\test_python_engine.py so no
' old Power Query fixture is ever reported as a current production PASS.
Public Sub RunWorkbookSelfTest()
    RunWorkbookSelfTestCore True
End Sub

Public Sub RunWorkbookSelfTestSilent()
    RunWorkbookSelfTestCore False
End Sub

Private Sub RunWorkbookSelfTestCore(ByVal showMessage As Boolean)
    Dim results As ListObject, passCount As Long, failCount As Long, infoCount As Long
    On Error GoTo EH
    Set results = ThisWorkbook.Worksheets("TEST_RESULTS").ListObjects("tblTestResults")
    If Not results.DataBodyRange Is Nothing Then results.DataBodyRange.Delete
    CheckTable results, "T01_TABLE_GR_DATA", "GR_DATA", "tblGrData", passCount, failCount
    CheckTable results, "T02_TABLE_INVOICES", "INVOICES", "tblInvoices", passCount, failCount
    CheckTable results, "T03_TABLE_VAT_LINES", "VAT_LINES", "tblVatLines", passCount, failCount
    CheckTable results, "T04_TABLE_LOG", "LOG", "tblLog", passCount, failCount
    CheckTable results, "T05_TABLE_PROFILES", "PARSER_PROFILES", "tblParserProfiles", passCount, failCount
    CheckPythonRuntime results, passCount, failCount
    CheckAsymmetricDateConfig results, passCount, failCount
    CheckSourceOwnership results, passCount, failCount, infoCount
    CheckNoteContract results, passCount, failCount, infoCount
    LogEvent "INFO", "RunWorkbookSelfTest", "SELF_TEST_DONE", CStr(passCount) & " passed; " & CStr(failCount) & " failed; " & CStr(infoCount) & " not-run batch checks.", "", "Review TEST_RESULTS; run Python regression tests before release."
    If showMessage Then MsgBox "Self-test completed: " & CStr(passCount) & " passed, " & CStr(failCount) & " failed, " & CStr(infoCount) & " chua co batch de kiem tra.", IIf(failCount = 0, vbInformation, vbExclamation)
    Exit Sub
EH:
    LogError "RunWorkbookSelfTest", "SELF_TEST_FAILED", Err.Description
    If showMessage Then MsgBox "Self-test failed. Review LOG.", vbCritical
End Sub

Private Sub CheckPythonRuntime(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    Dim launcher As String
    launcher = ThisWorkbook.Path & "\python\run_tool.bat"
    If Len(Dir$(launcher)) > 0 Then
        AddResult results, "T10_PYTHON_RUNTIME", "python\\run_tool.bat beside workbook", "Found", "PASS", "Excel will call the fixed Python entry point."
        passCount = passCount + 1
    Else
        AddResult results, "T10_PYTHON_RUNTIME", "python\\run_tool.bat beside workbook", "Missing", "FAIL", "Use the complete release folder; do not copy only the xlsm."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckAsymmetricDateConfig(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long)
    Dim config As ListObject, beforeValue As String, afterValue As String, i As Long
    Set config = ThisWorkbook.Worksheets("CONFIG").ListObjects("tblConfig")
    For i = 1 To config.ListRows.Count
        If CStr(config.DataBodyRange.Cells(i, 1).Value) = "GRBeforeInvoiceDays" Then beforeValue = CStr(config.DataBodyRange.Cells(i, 2).Value)
        If CStr(config.DataBodyRange.Cells(i, 1).Value) = "GRAfterInvoiceDays" Then afterValue = CStr(config.DataBodyRange.Cells(i, 2).Value)
    Next i
    If beforeValue = "5" And afterValue = "2" Then
        AddResult results, "T11_ASYMMETRIC_DATE", "GR before VAT=5; GR after VAT=2", beforeValue & "/" & afterValue, "PASS", "Python engine uses CONFIG; the old absolute ±2 rule is not used."
        passCount = passCount + 1
    Else
        AddResult results, "T11_ASYMMETRIC_DATE", "GR before VAT=5; GR after VAT=2", beforeValue & "/" & afterValue, "FAIL", "Review CONFIG before running a batch."
        failCount = failCount + 1
    End If
End Sub

Private Sub CheckSourceOwnership(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long, ByRef infoCount As Long)
    Dim allocations As ListObject, owners As Object, i As Long, sourceRow As String, invoiceId As String
    Set allocations = ThisWorkbook.Worksheets("ALLOCATIONS").ListObjects("tblAllocations")
    If allocations.DataBodyRange Is Nothing Then
        AddResult results, "T12_SOURCE_OWNERSHIP", "Run a Python batch", "NOT_RUN", "NOT_RUN", "No allocation rows in the current session."
        infoCount = infoCount + 1
        Exit Sub
    End If
    Set owners = CreateObject("Scripting.Dictionary")
    For i = 1 To allocations.ListRows.Count
        sourceRow = CStr(allocations.DataBodyRange.Cells(i, 5).Value)
        invoiceId = CStr(allocations.DataBodyRange.Cells(i, 2).Value)
        If owners.Exists(sourceRow) And CStr(owners(sourceRow)) <> invoiceId Then
            AddResult results, "T12_SOURCE_OWNERSHIP", "One SourceRow belongs to one invoice", sourceRow, "FAIL", "SourceRow is shared by two invoice IDs."
            failCount = failCount + 1
            Exit Sub
        End If
        owners(sourceRow) = invoiceId
    Next i
    AddResult results, "T12_SOURCE_OWNERSHIP", "One SourceRow belongs to one invoice", "0 shared SourceRows", "PASS", "Checks current Python batch allocation output."
    passCount = passCount + 1
End Sub

Private Sub CheckNoteContract(ByVal results As ListObject, ByRef passCount As Long, ByRef failCount As Long, ByRef infoCount As Long)
    Dim report As ListObject, i As Long, status As String, note As String
    Set report = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    If report.DataBodyRange Is Nothing Then
        AddResult results, "T13_PARTIAL_NOTE", "Run a Python batch", "NOT_RUN", "NOT_RUN", "No invoice report rows in the current session."
        infoCount = infoCount + 1
        Exit Sub
    End If
    For i = 1 To report.ListRows.Count
        status = UCase$(CStr(report.DataBodyRange.Cells(i, 8).Value))
        note = CStr(report.DataBodyRange.Cells(i, 11).Value)
        If status = "PARTIAL_MATCHED" And InStr(1, note, "Kh" & ChrW$(7899) & "p ", vbTextCompare) = 0 Then
            AddResult results, "T13_PARTIAL_NOTE", "Partial note states x/y material codes", note, "FAIL", "Python output note is missing its coverage statement."
            failCount = failCount + 1
            Exit Sub
        End If
    Next i
    AddResult results, "T13_PARTIAL_NOTE", "Partial note states x/y material codes", "Validated", "PASS", "Review exact unmatched-code names in BC_HOA_DON."
    passCount = passCount + 1
End Sub

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

Private Sub AddResult(ByVal results As ListObject, ByVal testId As String, ByVal expected As String, ByVal actual As String, ByVal result As String, ByVal evidence As String)
    Dim r As ListRow
    Set r = results.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 5).Value = Array(testId, expected, actual, result, evidence)
End Sub
