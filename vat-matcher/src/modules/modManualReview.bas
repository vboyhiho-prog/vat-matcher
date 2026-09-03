Attribute VB_Name = "modManualReview"
Option Explicit

Public Sub ApplyManualOverrides()
    Dim overrides As ListObject, report As ListObject, i As Long, reportRow As Long
    Dim invoiceNo As String, receiptSet As String, decision As String, reason As String
    On Error GoTo EH
    Set overrides = ThisWorkbook.Worksheets("MANUAL_OVERRIDES").ListObjects("tblManualOverrides")
    Set report = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    If overrides.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To overrides.ListRows.Count
        invoiceNo = NormalizeInvoiceNo(CStr(overrides.DataBodyRange.Cells(i, 1).Value))
        receiptSet = CStr(overrides.DataBodyRange.Cells(i, 2).Value)
        decision = UCase$(CStr(overrides.DataBodyRange.Cells(i, 3).Value))
        reason = CStr(overrides.DataBodyRange.Cells(i, 4).Value)
        If decision = "APPROVE" Or decision = "REJECT" Then
            If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 872, , "Manual override requires a reason: " & invoiceNo
            reportRow = FindReportRow(report, invoiceNo)
            If reportRow = 0 Then Err.Raise vbObjectError + 870, , "Invoice is not present in BC_HOA_DON: " & invoiceNo
            If decision = "APPROVE" And Not ReceiptSetExists(receiptSet) Then Err.Raise vbObjectError + 871, , "Receipt is not present in GR_DATA: " & receiptSet
            report.DataBodyRange.Cells(reportRow, 6).NumberFormat = "@"
            report.DataBodyRange.Cells(reportRow, 6).Value = receiptSet
            report.DataBodyRange.Cells(reportRow, 8).Value = "MANUAL_" & decision
            report.DataBodyRange.Cells(reportRow, 11).Value = reason
            overrides.DataBodyRange.Cells(i, 5).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
            LogEvent "INFO", "ApplyManualOverrides", "MANUAL_OVERRIDE_APPLIED", invoiceNo & " -> " & receiptSet & " / " & decision & " by " & Environ$("Username"), "", reason
        End If
    Next i
    Exit Sub
EH:
    LogError "ApplyManualOverrides", "MANUAL_OVERRIDE_FAILED", Err.Description
    MsgBox "Manual override failed. Review LOG.", vbCritical
End Sub

Public Sub CreateRenamePreviews()
    Dim report As ListObject, invoices As ListObject, pdfs As ListObject
    Dim i As Long, pdfId As String, vendor As String, receiptSet As String
    Dim vendors As Object, receipts As Object, proposedNames As Object, usedNames As Object, allOtherFactory As Object, baseName As String, candidateName As String, suffix As Long
    On Error GoTo EH
    Set report = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    If report.DataBodyRange Is Nothing Then Exit Sub
    Set vendors = CreateObject("Scripting.Dictionary")
    Set receipts = CreateObject("Scripting.Dictionary")
    Set proposedNames = CreateObject("Scripting.Dictionary")
    Set usedNames = CreateObject("Scripting.Dictionary")
    Set allOtherFactory = CreateObject("Scripting.Dictionary")
    For i = 1 To report.ListRows.Count
        pdfId = CStr(report.DataBodyRange.Cells(i, 2).Value)
        vendor = CStr(report.DataBodyRange.Cells(i, 4).Value)
        receiptSet = CStr(report.DataBodyRange.Cells(i, 6).Value)
        If Not vendors.Exists(pdfId) Then vendors.Add pdfId, vendor
        If Not allOtherFactory.Exists(pdfId) Then allOtherFactory.Add pdfId, True
        If UCase$(CStr(report.DataBodyRange.Cells(i, 8).Value)) <> "OTHER_FACTORY" Then allOtherFactory(pdfId) = False
        AddReceiptSet receipts, pdfId, receiptSet
    Next i
    For i = 1 To report.ListRows.Count
        pdfId = CStr(report.DataBodyRange.Cells(i, 2).Value)
        If Not proposedNames.Exists(pdfId) Then
            If CBool(allOtherFactory(pdfId)) Then
                baseName = "VAT XUONG KHAC"
            Else
                baseName = "VAT_" & SafeNamePart(CStr(vendors(pdfId))) & "_" & SafeNamePart(ReceiptText(receipts(pdfId)))
            End If
            candidateName = baseName & ".pdf"
            If usedNames.Exists(LCase$(candidateName)) Then
                candidateName = baseName & "_" & SafeNamePart(CStr(report.DataBodyRange.Cells(i, 3).Value)) & ".pdf"
                suffix = 2
                Do While usedNames.Exists(LCase$(candidateName))
                    candidateName = baseName & "_" & SafeNamePart(CStr(report.DataBodyRange.Cells(i, 3).Value)) & "_" & CStr(suffix) & ".pdf"
                    suffix = suffix + 1
                Loop
            End If
            proposedNames.Add pdfId, candidateName
            usedNames.Add LCase$(candidateName), True
        End If
        report.DataBodyRange.Cells(i, 9).Value = CStr(proposedNames(pdfId))
    Next i
    LogEvent "INFO", "CreateRenamePreviews", "RENAME_PREVIEW_READY", "Da tao ten de xuat VAT_NCC_SoPhieu; chua doi ten file.", "", "Nhap OK de doi ten, NG de bo qua."
    Exit Sub
EH:
    LogError "CreateRenamePreviews", "RENAME_PREVIEW_FAILED", Err.Description
    MsgBox "Tao ten de xuat that bai. Xem LOG.", vbCritical
End Sub

Private Sub AddReceiptSet(ByVal receipts As Object, ByVal pdfId As String, ByVal receiptSet As String)
    Dim receipt As Variant, uniqueReceipts As Object
    If Not receipts.Exists(pdfId) Then Set uniqueReceipts = CreateObject("Scripting.Dictionary"): receipts.Add pdfId, uniqueReceipts Else Set uniqueReceipts = receipts(pdfId)
    For Each receipt In Split(receiptSet, "+")
        If Len(Trim$(CStr(receipt))) > 0 Then uniqueReceipts(Trim$(CStr(receipt))) = True
    Next receipt
    If uniqueReceipts.Count = 0 Then uniqueReceipts("CHUA_CO_PHIEU") = True
End Sub

Private Function SafeNamePart(ByVal value As String) As String
    Dim invalidChars As Variant, i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars): value = Replace(value, CStr(invalidChars(i)), "_"): Next i
    SafeNamePart = Replace(Trim$(value), " ", "_")
End Function

Private Function ReceiptText(ByVal uniqueReceipts As Object) As String
    Dim receipt As Variant
    For Each receipt In uniqueReceipts.Keys
        If Len(ReceiptText) > 0 Then ReceiptText = ReceiptText & "+"
        ReceiptText = ReceiptText & CStr(receipt)
    Next receipt
End Function

Private Function FindReportRow(ByVal report As ListObject, ByVal invoiceNo As String) As Long
    Dim i As Long
    For i = 1 To report.ListRows.Count
        If NormalizeInvoiceNo(CStr(report.DataBodyRange.Cells(i, 3).Value)) = NormalizeInvoiceNo(invoiceNo) Then FindReportRow = i: Exit Function
    Next i
End Function

Private Function ReceiptSetExists(ByVal receiptSet As String) As Boolean
    Dim lo As ListObject, i As Long, receipt As Variant, found As Boolean
    Set lo = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    For Each receipt In Split(receiptSet, "+")
        found = False
        For i = 1 To lo.ListRows.Count
            If CStr(lo.DataBodyRange.Cells(i, 2).Value) = Trim$(CStr(receipt)) Then found = True: Exit For
        Next i
        If Not found Then Exit Function
    Next receipt
    ReceiptSetExists = (Len(Trim$(receiptSet)) > 0)
End Function

Private Function PdfIdForInvoice(ByVal invoices As ListObject, ByVal invoiceId As String) As String
    Dim i As Long
    For i = 1 To invoices.ListRows.Count
        If CStr(invoices.DataBodyRange.Cells(i, 1).Value) = invoiceId Then PdfIdForInvoice = CStr(invoices.DataBodyRange.Cells(i, 2).Value): Exit Function
    Next i
End Function

Private Function InvoiceCountForPdf(ByVal pdfs As ListObject, ByVal pdfId As String) As Long
    Dim i As Long
    For i = 1 To pdfs.ListRows.Count
        If CStr(pdfs.DataBodyRange.Cells(i, 1).Value) = pdfId Then InvoiceCountForPdf = CLng(Val(CStr(pdfs.DataBodyRange.Cells(i, 5).Value))): Exit Function
    Next i
End Function

Private Function NormalizeInvoiceNo(ByVal value As String) As String
    value = Trim$(value)
    If IsNumeric(value) Then NormalizeInvoiceNo = Right$("00000000" & CStr(CLng(Val(value))), 8) Else NormalizeInvoiceNo = value
End Function
