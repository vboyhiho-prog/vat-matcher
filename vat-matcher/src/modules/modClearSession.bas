Attribute VB_Name = "modClearSession"
Option Explicit

Public Sub ClearOldSessionData()
    Dim answer As VbMsgBoxResult, tableRefs As Variant, i As Long
    answer = MsgBox("Xoa du lieu phien cu: PDF, hoa don va ket qua khop? Du lieu file P dang nap va LOG lich su se duoc giu lai.", vbQuestion + vbYesNo, "Xac nhan xoa du lieu")
    If answer <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo EH
    tableRefs = Array( _
        Array("PDF_FILES", "tblPdfFiles"), Array("INVOICES", "tblInvoices"), Array("VAT_LINES", "tblVatLines"), _
        Array("MATCH_CANDIDATES", "tblCandidates"), Array("ALLOCATIONS", "tblAllocations"), _
        Array("BC_HOA_DON", "tblInvoiceReport"), Array("BC_PHIEU", "tblReceiptReport"), _
        Array("TEST_RESULTS", "tblTestResults"), Array("EMAIL_ATTACHMENTS", "tblEmailAttachments"), _
        Array("MANUAL_OVERRIDES", "tblManualOverrides"), Array("EMAIL_HINTS", "tblEmailHints"))
    For i = LBound(tableRefs) To UBound(tableRefs)
        ClearTableRows ThisWorkbook.Worksheets(CStr(tableRefs(i)(0))).ListObjects(CStr(tableRefs(i)(1)))
    Next i
    ResetPdfQueryForRelease
    ThisWorkbook.Worksheets("HOME").Range("B4").Value = ""
    UpdateDashboard
    Application.ScreenUpdating = True
    MsgBox "Da xoa du lieu phien cu. Du lieu P dang nap va LOG lich su duoc giu lai. LOG chi duoc xoa sau khi xuat goi chan doan thanh cong.", vbInformation
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Khong the xoa het du lieu. Xem LOG.", vbCritical
End Sub

Public Sub PrepareReleaseWorkbook()
    Dim tableRefs As Variant, i As Long
    Application.ScreenUpdating = False
    On Error GoTo EH
    tableRefs = Array( _
        Array("PDF_FILES", "tblPdfFiles"), Array("INVOICES", "tblInvoices"), Array("VAT_LINES", "tblVatLines"), _
        Array("GR_DATA", "tblGrData"), Array("MATCH_CANDIDATES", "tblCandidates"), Array("ALLOCATIONS", "tblAllocations"), _
        Array("BC_HOA_DON", "tblInvoiceReport"), Array("BC_PHIEU", "tblReceiptReport"), Array("LOG", "tblLog"), _
        Array("EMAIL_ATTACHMENTS", "tblEmailAttachments"), Array("MANUAL_OVERRIDES", "tblManualOverrides"), Array("EMAIL_HINTS", "tblEmailHints"))
    For i = LBound(tableRefs) To UBound(tableRefs)
        ClearTableRows ThisWorkbook.Worksheets(CStr(tableRefs(i)(0))).ListObjects(CStr(tableRefs(i)(1)))
    Next i
    ResetPdfQueryForRelease
    ClearAutoMaterialVendorRows
    ThisWorkbook.Worksheets("HOME").Range("B4").Value = ""
    ThisWorkbook.Worksheets("HOME").Range("D4").Value = ""
    ConfigureWorkbookUsability
    UpdateDashboard
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Err.Raise Err.Number, , "PrepareReleaseWorkbook: " & Err.Description
End Sub

Private Sub ClearAutoMaterialVendorRows()
    Dim lo As ListObject, i As Long
    Set lo = ThisWorkbook.Worksheets("MATERIAL_NCC_MAP").ListObjects("tblMaterialVendorMap")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = lo.ListRows.Count To 1 Step -1
        If UCase$(CStr(lo.DataBodyRange.Cells(i, 3).Value)) = "AUTO_TU_HOA_DON" Then lo.ListRows(i).Delete
    Next i
End Sub

Private Sub ClearTableRows(ByVal lo As ListObject)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
End Sub
