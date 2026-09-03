Attribute VB_Name = "modPdfParser"
Option Explicit

Public Sub ParsePdfRawToInvoices()
    Dim raw As ListObject, pdfs As ListObject, invoices As ListObject
    Dim rowIndex As Long, pageText As String, invoiceNo As String, taxCode As String
    Dim fileName As String, folderPath As String, pdfId As String, invoiceId As String
    Dim pageNo As Long, invoiceDate As Variant, vendor As String
    Dim fileIds As Object, invoiceCounts As Object, seenInvoices As Object
    On Error GoTo EH
    Set raw = ThisWorkbook.Worksheets("PQ_PDF_RAW").ListObjects("tblPdfRaw")
    Set pdfs = ThisWorkbook.Worksheets("PDF_FILES").ListObjects("tblPdfFiles")
    Set invoices = ThisWorkbook.Worksheets("INVOICES").ListObjects("tblInvoices")
    ClearTableRows pdfs
    ClearTableRows invoices
    Set fileIds = CreateObject("Scripting.Dictionary")
    Set invoiceCounts = CreateObject("Scripting.Dictionary")
    Set seenInvoices = CreateObject("Scripting.Dictionary")
    For rowIndex = 1 To raw.ListRows.Count
        If CStr(raw.DataBodyRange.Cells(rowIndex, raw.ListColumns("PdfTableKind").Index).Value) = "Page" Then
            pageText = CStr(raw.DataBodyRange.Cells(rowIndex, raw.ListColumns("PdfText").Index).Value)
            fileName = CStr(raw.DataBodyRange.Cells(rowIndex, raw.ListColumns("Name").Index).Value)
            folderPath = CStr(raw.DataBodyRange.Cells(rowIndex, raw.ListColumns("Folder Path").Index).Value)
            pdfId = EnsurePdfFile(pdfs, fileIds, invoiceCounts, folderPath, fileName)
            If IsPdfTextMissing(pageText) Then
                SetPdfStatus pdfs, pdfId, "NEEDS_OCR"
                GoTo NextRawRow
            End If
            taxCode = ExtractTaxCode(pageText)
            invoiceNo = ExtractInvoiceNo(pageText, taxCode)
            If Len(invoiceNo) > 0 Then
                pageNo = PageNumber(CStr(raw.DataBodyRange.Cells(rowIndex, raw.ListColumns("PdfTableId").Index).Value))
                vendor = VendorForTaxCode(taxCode)
                invoiceDate = ExtractInvoiceDate(pageText, taxCode)
                invoiceId = "INV-" & invoiceNo
                If seenInvoices.Exists(pdfId & "|" & invoiceNo) Then
                    invoices.DataBodyRange.Cells(CLng(seenInvoices(pdfId & "|" & invoiceNo)), 4).Value = pageNo
                Else
                    AddInvoice invoices, invoiceId, pdfId, pageNo, pageNo, invoiceNo, invoiceDate, vendor, taxCode, "PARSED"
                    seenInvoices.Add pdfId & "|" & invoiceNo, invoices.ListRows.Count
                    invoiceCounts(pdfId) = CLng(invoiceCounts(pdfId)) + 1
                End If
            End If
        End If
NextRawRow:
    Next rowIndex
    UpdatePdfInvoiceCounts pdfs, invoiceCounts
    LogEvent "INFO", "ParsePdfRawToInvoices", "PDF_INVOICE_PARSE_OK", CStr(invoices.ListRows.Count) & " invoice records created.", "", "Run matching only after reviewing parsed invoices."
    Exit Sub
EH:
    LogError "ParsePdfRawToInvoices", "PDF_INVOICE_PARSE_FAILED", Err.Description
    MsgBox "PDF invoice parsing failed. Review LOG.", vbCritical
End Sub

Public Function IsPdfTextMissing(ByVal pageText As String) As Boolean
    IsPdfTextMissing = (Len(Trim$(pageText)) = 0)
End Function

Public Sub ParseVatLinesFromPdfRaw()
    Dim raw As ListObject, lines As ListObject, rowIndex As Long, seq As Long
    Dim pageText As String, invoiceNo As String, invoiceId As String, seqByInvoice As Object, taxCode As String, linePattern As String
    Dim re As Object, reCodeFirst As Object, reNumberFirst As Object, reProfile As Object, matches As Object, m As Object, seenLines As Object
    On Error GoTo EH
    Set raw = ThisWorkbook.Worksheets("PQ_PDF_RAW").ListObjects("tblPdfRaw")
    Set lines = ThisWorkbook.Worksheets("VAT_LINES").ListObjects("tblVatLines")
    ClearTableRows lines
    Set seqByInvoice = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "Linh ki.n\s+([A-Z0-9]+)[^|]*\|\s*[^|]*\|\s*([0-9.]+,[0-9]+)"
    Set reCodeFirst = CreateObject("VBScript.RegExp")
    reCodeFirst.Global = True
    reCodeFirst.IgnoreCase = True
    reCodeFirst.Pattern = "([A-Z]{3}[A-Z0-9]*[0-9][A-Z0-9]*)[\s\S]{0,220}?\|\s*[0-9]{1,2}\s*\|\s*[^|]+\|\s*([0-9.]+,[0-9]+)"
    Set reNumberFirst = CreateObject("VBScript.RegExp")
    reNumberFirst.Global = True
    reNumberFirst.IgnoreCase = True
    reNumberFirst.Pattern = "\|\s*[0-9]{1,2}\s*\|\s*([A-Z]{3}[A-Z0-9]*[0-9][A-Z0-9]*)[^|]*\|\s*[^|]+\|\s*([0-9.]+,[0-9]+)"
    Set reProfile = CreateObject("VBScript.RegExp")
    reProfile.Global = True: reProfile.IgnoreCase = True
    Set seenLines = CreateObject("Scripting.Dictionary")
    For rowIndex = 1 To raw.ListRows.Count
        If CStr(raw.DataBodyRange.Cells(rowIndex, raw.ListColumns("PdfTableKind").Index).Value) = "Page" Then
            pageText = CStr(raw.DataBodyRange.Cells(rowIndex, raw.ListColumns("PdfText").Index).Value)
            taxCode = ExtractTaxCode(pageText)
            invoiceNo = ExtractInvoiceNo(pageText, taxCode)
            If Len(invoiceNo) > 0 Then
                invoiceId = "INV-" & invoiceNo
                If seqByInvoice.Exists(invoiceId) Then seq = CLng(seqByInvoice(invoiceId)) Else seq = 0
                Set matches = re.Execute(pageText)
                For Each m In matches
                    AddPdfVatLine lines, seenLines, invoiceNo, invoiceId, seq, CStr(m.SubMatches(0)), CStr(m.SubMatches(1))
                Next m
                linePattern = ProfilePattern(taxCode, 6)
                If Len(linePattern) > 0 Then
                    reProfile.Pattern = linePattern
                    Set matches = reProfile.Execute(pageText)
                    For Each m In matches
                        If m.SubMatches.Count >= 2 Then AddPdfVatLine lines, seenLines, invoiceNo, invoiceId, seq, CStr(m.SubMatches(0)), CStr(m.SubMatches(1))
                    Next m
                End If
                If InStr(1, pageText, "2500645835", vbTextCompare) > 0 Then
                    Set matches = reCodeFirst.Execute(pageText)
                    For Each m In matches
                        AddPdfVatLine lines, seenLines, invoiceNo, invoiceId, seq, CStr(m.SubMatches(0)), CStr(m.SubMatches(1))
                    Next m
                    Set matches = reNumberFirst.Execute(pageText)
                    For Each m In matches
                        AddPdfVatLine lines, seenLines, invoiceNo, invoiceId, seq, CStr(m.SubMatches(0)), CStr(m.SubMatches(1))
                    Next m
                End If
                seqByInvoice(invoiceId) = seq
            End If
        End If
    Next rowIndex
    LogEvent "INFO", "ParseVatLinesFromPdfRaw", "PDF_VAT_LINES_OK", CStr(lines.ListRows.Count) & " VAT lines created.", "", "Review VAT_LINES before matching."
    Exit Sub
EH:
    LogError "ParseVatLinesFromPdfRaw", "PDF_VAT_LINES_FAILED", Err.Description
    MsgBox "VAT line parsing failed. Review LOG.", vbCritical
End Sub

Private Sub AddPdfVatLine(ByVal lines As ListObject, ByVal seenLines As Object, ByVal invoiceNo As String, ByVal invoiceId As String, ByRef seq As Long, ByVal material As String, ByVal rawQty As String)
    Dim lineKey As String, r As ListRow
    lineKey = invoiceId & "|" & material & "|" & rawQty
    If seenLines.Exists(lineKey) Then Exit Sub
    seenLines.Add lineKey, True
    seq = seq + 1
    Set r = lines.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 9).Value = Array("VL-" & invoiceNo & "-" & Format$(seq, "000"), invoiceId, seq, material, NormalizeMaterial(material), rawQty, QtyFromPdf(rawQty), "EA", 0.9)
End Sub

Private Function QtyFromPdf(ByVal rawQty As String) As Double
    QtyFromPdf = Val(Replace(Replace(rawQty, ".", ""), ",", "."))
End Function

Private Function EnsurePdfFile(ByVal pdfs As ListObject, ByVal fileIds As Object, ByVal invoiceCounts As Object, ByVal folderPath As String, ByVal fileName As String) As String
    Dim key As String, r As ListRow
    key = folderPath & fileName
    If Not fileIds.Exists(key) Then
        EnsurePdfFile = "PDF-" & Format$(fileIds.Count + 1, "0000")
        fileIds.Add key, EnsurePdfFile
        invoiceCounts.Add EnsurePdfFile, 0
        Set r = pdfs.ListRows.Add
        r.Range.Cells(1, 1).Resize(1, 5).Value = Array(EnsurePdfFile, key, "", "POWER_QUERY_OK", 0)
    Else
        EnsurePdfFile = CStr(fileIds(key))
    End If
End Function

Private Sub UpdatePdfInvoiceCounts(ByVal pdfs As ListObject, ByVal invoiceCounts As Object)
    Dim i As Long, pdfId As String
    For i = 1 To pdfs.ListRows.Count
        pdfId = CStr(pdfs.DataBodyRange.Cells(i, 1).Value)
        pdfs.DataBodyRange.Cells(i, 5).Value = CLng(invoiceCounts(pdfId))
    Next i
End Sub

Private Sub SetPdfStatus(ByVal pdfs As ListObject, ByVal pdfId As String, ByVal status As String)
    Dim i As Long
    For i = 1 To pdfs.ListRows.Count
        If CStr(pdfs.DataBodyRange.Cells(i, 1).Value) = pdfId Then
            pdfs.DataBodyRange.Cells(i, 4).Value = status
            Exit Sub
        End If
    Next i
End Sub

Private Sub AddInvoice(ByVal invoices As ListObject, ByVal invoiceId As String, ByVal pdfId As String, ByVal pageFrom As Long, ByVal pageTo As Long, ByVal invoiceNo As String, ByVal invoiceDate As Variant, ByVal vendor As String, ByVal taxCode As String, ByVal parseStatus As String)
    Dim r As ListRow
    Set r = invoices.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 9).Value = Array(invoiceId, pdfId, pageFrom, pageTo, "", invoiceDate, vendor, "", parseStatus)
    r.Range.Cells(1, 5).NumberFormat = "@"
    r.Range.Cells(1, 5).Value = invoiceNo
    r.Range.Cells(1, 8).NumberFormat = "@"
    r.Range.Cells(1, 8).Value = taxCode
End Sub

Private Sub ClearTableRows(ByVal lo As ListObject)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
End Sub

Private Function ExtractInvoiceNo(ByVal pageText As String, Optional ByVal taxCode As String = "") As String
    Dim configuredPattern As String
    configuredPattern = ProfilePattern(taxCode, 4)
    If Len(configuredPattern) > 0 Then ExtractInvoiceNo = RegexFirst(pageText, configuredPattern)
    If Len(ExtractInvoiceNo) = 0 Then ExtractInvoiceNo = RegexFirst(pageText, "\b0{4}[0-9]{4}\b")
End Function

Private Function ExtractTaxCode(ByVal pageText As String) As String
    ExtractTaxCode = RegexFirst(pageText, "\b[0-9]{10}\b")
End Function

Private Function ExtractInvoiceDate(ByVal pageText As String, Optional ByVal taxCode As String = "") As Variant
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    If Len(ProfilePattern(taxCode, 5)) > 0 Then
        re.Pattern = ProfilePattern(taxCode, 5)
    Else
        re.Pattern = "([0-3]?[0-9])\s*\|\s*th.ng\s*\|\s*([01]?[0-9])\s*\|\s*n.m\s*\|\s*(20[0-9]{2})"
    End If
    Set matches = re.Execute(pageText)
    If matches.Count = 0 And Len(ProfilePattern(taxCode, 5)) = 0 Then
        re.Pattern = "Ng.y\s*\(Date\)\s*([0-3]?[0-9])\s*th.ng\s*\(month\)\s*([01]?[0-9])\s*n.m\s*\(year\)\s*(20[0-9]{2})"
        Set matches = re.Execute(pageText)
    End If
    If matches.Count > 0 Then ExtractInvoiceDate = DateSerial(CInt(matches(0).SubMatches(2)), CInt(matches(0).SubMatches(1)), CInt(matches(0).SubMatches(0)))
End Function

Private Function ProfilePattern(ByVal taxCode As String, ByVal columnIndex As Long) As String
    Dim profiles As ListObject, i As Long
    If Len(taxCode) = 0 Then Exit Function
    Set profiles = ThisWorkbook.Worksheets("PARSER_PROFILES").ListObjects("tblParserProfiles")
    If profiles.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To profiles.ListRows.Count
        If CStr(profiles.DataBodyRange.Cells(i, 3).Value) = taxCode Then
            ProfilePattern = Trim$(CStr(profiles.DataBodyRange.Cells(i, columnIndex).Value))
            Exit Function
        End If
    Next i
End Function

Private Function RegexFirst(ByVal value As String, ByVal pattern As String) As String
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = pattern
    Set matches = re.Execute(value)
    If matches.Count > 0 Then RegexFirst = CStr(matches(0).Value)
End Function

Private Function VendorForTaxCode(ByVal taxCode As String) As String
    Dim lo As ListObject, i As Long
    Set lo = ThisWorkbook.Worksheets("NCC_MAP").ListObjects("tblVendorMap")
    For i = 1 To lo.ListRows.Count
        If CDbl(Val(CStr(lo.DataBodyRange.Cells(i, 2).Value))) = CDbl(Val(taxCode)) Then
            VendorForTaxCode = CStr(lo.DataBodyRange.Cells(i, 1).Value)
            Exit Function
        End If
    Next i
    VendorForTaxCode = "UNMAPPED"
End Function

Private Function PageNumber(ByVal pageId As String) As Long
    PageNumber = CLng(Val(Right$(pageId, 3)))
End Function
