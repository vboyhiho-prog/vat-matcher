Attribute VB_Name = "modBootstrap"
Option Explicit

Public Sub SetupWorkbook()
    Dim specs As Variant, i As Long
    specs = Array( _
        Array("HOME", "tblHome", Array("Key", "Value", "Status")), _
        Array("CONFIG", "tblConfig", Array("Key", "Value", "Description")), _
        Array("NCC_MAP", "tblVendorMap", Array("CanonicalVendor", "TaxCode", "RawAlias", "FileAlias", "ParserProfile", "Active")), _
        Array("PARSER_PROFILES", "tblParserProfiles", Array("ProfileID", "CanonicalVendor", "TaxCode", "InvoiceNoPattern", "DatePattern", "LinePattern", "Status", "SamplePdf", "Notes")), _
        Array("NCC_CONFIG_IMPORT", "tblVendorConfigImport", Array("CanonicalVendor", "TaxCode", "RawAlias", "FileAlias", "ProfileID", "InvoiceNoPattern", "DatePattern", "LinePattern", "ProfileStatus", "Active", "Notes")), _
        Array("PDF_FILES", "tblPdfFiles", Array("PdfID", "OriginalPath", "Fingerprint", "ExtractStatus", "InvoiceCount")), _
        Array("EMAIL_ATTACHMENTS", "tblEmailAttachments", Array("MsgPath", "Subject", "AttachmentName", "AttachmentType", "ExtractStatus", "SavedPath")), _
        Array("EMAIL_HINTS", "tblEmailHints", Array("MsgPath", "IB", "Material", "Qty", "GRDate", "VendorHint", "Confidence", "RawContext")), _
        Array("PQ_PDF_RAW", "tblPdfRaw", Array("Name", "Extension", "Folder Path", "Date modified", "PdfTableId", "PdfTableKind", "PdfTableName", "PdfTableData", "PdfText")), _
        Array("INVOICES", "tblInvoices", Array("InvoiceID", "PdfID", "PageFrom", "PageTo", "InvoiceNo", "InvoiceDate", "VendorCanonical", "TaxCode", "ParseStatus")), _
        Array("VAT_LINES", "tblVatLines", Array("VatLineID", "InvoiceID", "Seq", "MaterialRaw", "MaterialNorm", "QtyRaw", "Qty", "Unit", "ParseConfidence", "ScopeStatus")), _
        Array("MATERIAL_NCC_MAP", "tblMaterialVendorMap", Array("MaterialNorm", "CanonicalVendor", "Source", "Active", "Note")), _
        Array("MATERIAL_SCOPE_MAP", "tblMaterialScopeMap", Array("MaterialNorm", "ScopeStatus", "Note")), _
        Array("GR_DATA", "tblGrData", Array("SourceRow", "ReceiptNo", "Material", "Vendor", "QtyDoc", "QtyActual", "QtyMatch", "IB", "GRDate", "Flags")), _
        Array("MATCH_CANDIDATES", "tblCandidates", Array("CandidateID", "InvoiceID", "ReceiptSet", "VendorScore", "MaterialScore", "QtyScore", "DateScore", "HintScore", "TotalScore", "Rank", "Reasons")), _
        Array("ALLOCATIONS", "tblAllocations", Array("AllocationID", "InvoiceID", "VatLineID", "ReceiptNo", "SourceRow", "AllocatedQty", "Residual", "Status")), _
        Array("MANUAL_OVERRIDES", "tblManualOverrides", Array("InvoiceNo", "ReceiptSet", "Decision", "Reason", "UpdatedAt")), _
        Array("MATCH_HINTS", "tblMatchHints", Array("InvoiceNo", "ReceiptSet", "HintType", "Active", "Note")), _
        Array("BC_HOA_DON", "tblInvoiceReport", Array("RunID", "PDF", "InvoiceNo", "Vendor", "InvoiceDate", "ProposedReceipts", "TotalScore", "Status", "RenamePreview", "Decision", "Note")), _
        Array("BC_PHIEU", "tblReceiptReport", Array("ReceiptNo", "Vendor", "GRDate", "LinkedInvoices", "Coverage", "Status")), _
        Array("LOG", "tblLog", Array("Timestamp", "RunID", "Severity", "Module", "Procedure", "Code", "Message", "OriginalPath", "NewPath")), _
        Array("TEST_RESULTS", "tblTestResults", Array("TestID", "Expected", "Actual", "Result", "Evidence")))
    Application.ScreenUpdating = False
    On Error GoTo EH
    For i = LBound(specs) To UBound(specs)
        EnsureSheetAndTable CStr(specs(i)(0)), CStr(specs(i)(1)), specs(i)(2)
    Next i
    NormalizeHomeSheet
    SeedConfig
    SeedVendorMap
    SeedMatchHints
    SeedMaterialScopeMap
    SeedMaterialVendorMap
    SeedVendorConfigTemplate
    ConfigureWorkbookUsability
    Worksheets("MANUAL_OVERRIDES").ListObjects("tblManualOverrides").ListColumns(1).Range.NumberFormat = "@"
    Worksheets("MANUAL_OVERRIDES").ListObjects("tblManualOverrides").ListColumns(2).Range.NumberFormat = "@"
    Worksheets("HOME").Range("A1:C1").Value = Array("VAT Matcher - Power Query batch PDF", "", "")
    Worksheets("HOME").Columns("A:C").AutoFit
CleanUp:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    MsgBox "Setup failed: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Private Sub NormalizeHomeSheet()
    Dim ws As Worksheet, lo As ListObject, pdfPath As String, trackingPath As String, legacyValue As String
    Set ws = Worksheets("HOME")
    pdfPath = Trim$(CStr(ws.Range("B4").Value))
    trackingPath = Trim$(CStr(ws.Range("D4").Value))
    If Len(pdfPath) = 0 Then
        legacyValue = Trim$(CStr(ws.Range("B2").Value))
        If InStr(1, legacyValue, "\", vbTextCompare) > 0 Then pdfPath = legacyValue
    End If
    If Len(trackingPath) = 0 Then trackingPath = Trim$(CStr(ws.Range("D3").Value))
    On Error Resume Next
    ws.ListObjects("tblHome").Delete
    On Error GoTo 0
    ws.Range("A2:D5").ClearContents
    ws.Range("A2:C2").Value = Array("Key", "Value", "Status")
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A2:C2"), , xlYes)
    lo.Name = "tblHome"
    lo.TableStyle = "TableStyleMedium2"
    ws.Range("A4").Value = "PDF_FOLDER_PATH"
    ws.Range("B4").Value = pdfPath
    ws.Range("C4").Value = "TRACKING_P_PATH"
    ws.Range("D4").Value = trackingPath
End Sub

Private Sub SeedMaterialScopeMap()
    Dim lo As ListObject, ws As Worksheet
    Set ws = Worksheets("MATERIAL_SCOPE_MAP")
    Set lo = ws.ListObjects("tblMaterialScopeMap")
    ws.Range("A1").Value = "Add only confirmed external material codes as OUT_OF_SCOPE_MATERIAL. Unknown codes remain reviewable and are never ignored automatically."
End Sub

Private Sub SeedMaterialVendorMap()
    Dim ws As Worksheet
    Set ws = Worksheets("MATERIAL_NCC_MAP")
    ws.Range("A1").Value = "Tool tu dong gan ma vat tu vao NCC theo hoa don da nhan dien MST. Khong dung cot ten NCC tu file P. Chi sua dong nay khi can chinh sua mapping ma vat tu."
End Sub

Private Sub EnsureSheetAndTable(ByVal sheetName As String, ByVal tableName As String, ByVal headers As Variant)
    Dim ws As Worksheet, lo As ListObject, i As Long
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count)): ws.Name = sheetName
    On Error Resume Next: Set lo = ws.ListObjects(tableName): On Error GoTo 0
    If lo Is Nothing Then
        For i = LBound(headers) To UBound(headers): ws.Cells(2, i + 1).Value = headers(i): Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(2, 1), ws.Cells(2, UBound(headers) + 1)), , xlYes)
        lo.Name = tableName
        lo.TableStyle = "TableStyleMedium2"
    Else
        For i = LBound(headers) To UBound(headers)
            If Not TableHasColumn(lo, CStr(headers(i))) Then lo.ListColumns.Add.Name = CStr(headers(i))
        Next i
    End If
End Sub

Private Function TableHasColumn(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    Dim col As ListColumn
    For Each col In lo.ListColumns
        If col.Name = columnName Then TableHasColumn = True: Exit Function
    Next col
End Function

Private Sub SeedConfig()
    Dim lo As ListObject: Set lo = Worksheets("CONFIG").ListObjects("tblConfig")
    EnsureConfigRow lo, "DateWindowDays", 2, "Allowed invoice/GR date difference"
    EnsureConfigRow lo, "PdfIngestion", "Power Query Pdf.Tables", "No external executable"
    EnsureConfigRow lo, "QtySource", "QTY_DOCUMENT", "QTY_DOCUMENT by default; QTY_ACTUAL is optional. Mismatch rows are review-only."
End Sub

Private Sub EnsureConfigRow(ByVal lo As ListObject, ByVal keyName As String, ByVal keyValue As Variant, ByVal description As String)
    Dim i As Long, r As ListRow
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(i, 1).Value)), keyName, vbTextCompare) = 0 Then
            lo.DataBodyRange.Cells(i, 2).Value = keyValue
            lo.DataBodyRange.Cells(i, 3).Value = description
            Exit Sub
        End If
    Next i
    Set r = lo.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 3).Value = Array(keyName, keyValue, description)
End Sub

Public Sub ConfigureWorkbookUsability()
    Dim ws As Worksheet, separator As String
    On Error GoTo SafeExit
    separator = Application.International(xlListSeparator)
    Set ws = ThisWorkbook.Worksheets("BC_HOA_DON")
    ws.Columns("A").ColumnWidth = 21
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C:D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 13
    ws.Columns("F").ColumnWidth = 24
    ws.Columns("G:H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 38
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 70
    ws.Columns("E").NumberFormat = "dd/mm/yyyy"
    ws.Columns("G").NumberFormat = "0.0"
    ws.Columns("K").WrapText = True
    With ws.Range("J3:J5000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="OK" & separator & "NG"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Set ws = ThisWorkbook.Worksheets("BC_PHIEU")
    ws.Columns("C").NumberFormat = "dd/mm/yyyy"
    ws.Columns("A:F").AutoFit
    Set ws = ThisWorkbook.Worksheets("GR_DATA")
    ws.Columns("I").NumberFormat = "dd/mm/yyyy"
    ws.Columns("A:B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E:G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 13
    ws.Columns("J").ColumnWidth = 28
    Set ws = ThisWorkbook.Worksheets("LOG")
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B:F").ColumnWidth = 18
    ws.Columns("G:I").ColumnWidth = 45
    ws.Columns("G:I").WrapText = True
SafeExit:
End Sub

Private Sub SeedVendorMap()
    Dim lo As ListObject: Set lo = Worksheets("NCC_MAP").ListObjects("tblVendorMap")
    If lo.ListRows.Count > 0 Then Exit Sub
    AddRow lo, Array("THANH_DAT", "0106097880", "THANH DAT", "THANH_DAT", "THANH_DAT", True), Array("LTV", "2500645835", "CONG TY TNHH CONG NGHIEP LTV", "LTV", "LTV", True)
End Sub

Private Sub SeedMatchHints()
    Dim lo As ListObject: Set lo = Worksheets("MATCH_HINTS").ListObjects("tblMatchHints")
    If lo.ListRows.Count > 0 Then Exit Sub
    AddRow lo, Array("00002212", "4983+5005+5007", "REFERENCE", True, "Verified LTV reference; review before production use.")
    lo.ListColumns(1).Range.NumberFormat = "@"
    lo.ListColumns(2).Range.NumberFormat = "@"
End Sub

Private Sub SeedVendorConfigTemplate()
    Dim lo As ListObject, ws As Worksheet
    Set ws = Worksheets("NCC_CONFIG_IMPORT")
    Set lo = ws.ListObjects("tblVendorConfigImport")
    ws.Range("A1").Value = "Paste supplier configuration below, then run LoadVendorProfilesFromConfig. EXAMPLE rows are ignored."
    If lo.ListRows.Count > 0 Then Exit Sub
    AddRow lo, Array("EXAMPLE_VENDOR", "0123456789", "EXAMPLE SUPPLIER", "EXAMPLE", "EXAMPLE_PROFILE", "\\b0{4}[0-9]{4}\\b", "DATE_PATTERN", "LINE_PATTERN", "DRAFT", False, "Replace this row with a real supplier.")
    lo.ListColumns(1).Range.NumberFormat = "@"
    lo.ListColumns(2).Range.NumberFormat = "@"
End Sub

Private Sub AddRow(ByVal lo As ListObject, ParamArray values())
    Dim i As Long, r As ListRow
    For i = LBound(values) To UBound(values)
        Set r = lo.ListRows.Add
        r.Range.Cells(1, 1).Resize(1, UBound(values(i)) + 1).Value = values(i)
    Next i
End Sub

Private Sub WriteHome(ByVal k As String, ByVal v As String, ByVal s As String)
    Dim lo As ListObject, r As ListRow: Set lo = Worksheets("HOME").ListObjects("tblHome")
    Set r = lo.ListRows.Add: r.Range.Cells(1, 1).Resize(1, 3).Value = Array(k, v, s)
End Sub
