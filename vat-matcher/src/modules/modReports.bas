Attribute VB_Name = "modReports"
Option Explicit

Public Sub BuildReceiptReport()
    Dim gr As ListObject, allocations As ListObject, report As ListObject, invoiceReport As ListObject
    Dim receiptInfo As Object, linkedInvoices As Object, invoiceStatuses As Object, i As Long, receipt As String, receiptKey As Variant, info As Variant, r As ListRow
    Dim invoiceId As Variant, receiptStatus As String, grData As Variant, allocationData As Variant, invoiceData As Variant
    On Error GoTo EH
    Set gr = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    Set allocations = ThisWorkbook.Worksheets("ALLOCATIONS").ListObjects("tblAllocations")
    Set report = ThisWorkbook.Worksheets("BC_PHIEU").ListObjects("tblReceiptReport")
    Set invoiceReport = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    Set receiptInfo = CreateObject("Scripting.Dictionary")
    Set linkedInvoices = CreateObject("Scripting.Dictionary")
    Set invoiceStatuses = CreateObject("Scripting.Dictionary")
    If Not report.DataBodyRange Is Nothing Then report.DataBodyRange.Delete
    If Not gr.DataBodyRange Is Nothing Then
        grData = gr.DataBodyRange.Value2
        For i = 1 To UBound(grData, 1)
            receipt = CStr(grData(i, 2))
            If Len(receipt) > 0 And Not receiptInfo.Exists(receipt) Then receiptInfo.Add receipt, Array(CStr(grData(i, 4)), grData(i, 9))
        Next i
    End If
    If Not allocations.DataBodyRange Is Nothing Then
        allocationData = allocations.DataBodyRange.Value2
        For i = 1 To UBound(allocationData, 1)
            receipt = CStr(allocationData(i, 4))
            If Len(receipt) > 0 Then
                If Not linkedInvoices.Exists(receipt) Then linkedInvoices.Add receipt, CreateObject("Scripting.Dictionary")
                linkedInvoices(receipt)(CStr(allocationData(i, 2))) = True
            End If
        Next i
    End If
    If Not invoiceReport.DataBodyRange Is Nothing Then
        invoiceData = invoiceReport.DataBodyRange.Value2
        For i = 1 To UBound(invoiceData, 1)
            invoiceStatuses("INV-" & CStr(invoiceData(i, 3))) = UCase$(CStr(invoiceData(i, 8)))
        Next i
    End If
    For Each receiptKey In linkedInvoices.Keys
        receipt = CStr(receiptKey)
        info = receiptInfo(receipt)
        receiptStatus = "MATCHED"
        For Each invoiceId In linkedInvoices(receipt).Keys
            If Not invoiceStatuses.Exists(CStr(invoiceId)) Or CStr(invoiceStatuses(invoiceId)) <> "MATCHED" Then receiptStatus = "REVIEW"
        Next invoiceId
        Set r = report.ListRows.Add
        r.Range.Cells(1, 1).Resize(1, 6).Value = Array(receipt, info(0), info(1), Join(linkedInvoices(receipt).Keys, ","), linkedInvoices(receipt).Count, receiptStatus)
    Next receiptKey
    LogEvent "INFO", "BuildReceiptReport", "RECEIPT_REPORT_OK", CStr(report.ListRows.Count) & " receipt rows built.", "", "Review allocations and coverage."
    Exit Sub
EH:
    LogError "BuildReceiptReport", "RECEIPT_REPORT_FAILED", Err.Description
End Sub
