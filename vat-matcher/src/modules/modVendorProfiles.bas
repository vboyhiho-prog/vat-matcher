Attribute VB_Name = "modVendorProfiles"
Option Explicit

Public Sub LoadVendorProfilesFromConfig()
    Dim source As ListObject, vendorMap As ListObject, profiles As ListObject
    Dim i As Long, vendor As String, taxCode As String, profileId As String, status As String
    On Error GoTo EH
    Set source = ThisWorkbook.Worksheets("NCC_CONFIG_IMPORT").ListObjects("tblVendorConfigImport")
    Set vendorMap = ThisWorkbook.Worksheets("NCC_MAP").ListObjects("tblVendorMap")
    Set profiles = ThisWorkbook.Worksheets("PARSER_PROFILES").ListObjects("tblParserProfiles")
    If source.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To source.ListRows.Count
        vendor = Trim$(CStr(source.DataBodyRange.Cells(i, 1).Value))
        taxCode = Trim$(CStr(source.DataBodyRange.Cells(i, 2).Value))
        profileId = Trim$(CStr(source.DataBodyRange.Cells(i, 5).Value))
        status = UCase$(Trim$(CStr(source.DataBodyRange.Cells(i, 9).Value)))
        If Len(vendor) > 0 And Left$(vendor, 8) <> "EXAMPLE_" Then
            If Len(profileId) = 0 Then GoTo NextVendor
            If Len(taxCode) = 0 Then Err.Raise vbObjectError + 910, , "TaxCode is required when creating a parser profile: " & vendor
            If status <> "DRAFT" And status <> "ACTIVE" Then Err.Raise vbObjectError + 911, , "ProfileStatus must be DRAFT or ACTIVE: " & vendor
            UpsertVendorMap vendorMap, source, i
            UpsertParserProfile profiles, source, i
            LogEvent "INFO", "LoadVendorProfilesFromConfig", "VENDOR_PROFILE_LOADED", vendor & " / " & profileId & " / " & status, "", "Keep DRAFT until a sample PDF test passes."
        End If
NextVendor:
    Next i
    MsgBox "Vendor configuration loaded. Review PARSER_PROFILES before activation.", vbInformation
    Exit Sub
EH:
    LogError "LoadVendorProfilesFromConfig", "VENDOR_PROFILE_LOAD_FAILED", Err.Description
    MsgBox "Vendor configuration failed. Review LOG.", vbCritical
End Sub

Public Sub TaoDanhSachNCCTuFileP()
    Dim gr As ListObject, vendorMap As ListObject, config As ListObject, i As Long, rawVendor As String, canonical As String, added As Long
    On Error GoTo EH
    Set gr = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    If gr.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 920, , "Hay nap file P truoc."
    Set vendorMap = ThisWorkbook.Worksheets("NCC_MAP").ListObjects("tblVendorMap")
    Set config = ThisWorkbook.Worksheets("NCC_CONFIG_IMPORT").ListObjects("tblVendorConfigImport")
    For i = 1 To gr.ListRows.Count
        rawVendor = Trim$(CStr(gr.DataBodyRange.Cells(i, 4).Value))
        If Len(rawVendor) > 0 And Not VendorAliasExists(vendorMap, rawVendor) Then
            canonical = NormalizeText(rawVendor)
            AddVendorMapRow vendorMap, canonical, rawVendor
            AddVendorConfigRow config, canonical, rawVendor
            added = added + 1
        End If
    Next i
    ThisWorkbook.Worksheets("NCC_CONFIG_IMPORT").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("NCC_CONFIG_IMPORT").Activate
    MsgBox "Da tao " & CStr(added) & " NCC tu file P. NCC_MAP da duoc cap nhat. Neu ten NCC tren hoa don khac ten trong P, them ten do vao cot RawAlias cua NCC_MAP.", vbInformation
    Exit Sub
EH:
    LogError "TaoDanhSachNCCTuFileP", "VENDOR_FROM_GR_FAILED", Err.Description
    MsgBox "Khong the tao danh sach NCC. Hay nap file P truoc.", vbCritical
End Sub

Public Sub CauHinhNCCVaParser()
    Dim vendor As String, taxCode As String, rawAlias As String, samplePdf As String, profileId As String
    Dim vendorMap As ListObject, profiles As ListObject, r As ListRow
    vendor = Trim$(InputBox("Nhap ten viet tat dung de khop voi file P, vi du: LTV", "Cau hinh NCC + parser"))
    If Len(vendor) = 0 Then Exit Sub
    taxCode = Trim$(InputBox("Nhap DUNG ma so thue tren hoa don PDF. Day la khoa de tool nhan dien NCC, vi du: 2500645835", "Cau hinh NCC + parser"))
    If Len(taxCode) <> 10 Or Not IsNumeric(taxCode) Then MsgBox "Ma so thue phai gom dung 10 chu so tren hoa don.", vbExclamation: Exit Sub
    rawAlias = Trim$(InputBox("Nhap ten NCC dung nhu tren hoa don (chi de tham khao), vi du: LTV", "Cau hinh NCC + parser", vendor))
    If Len(rawAlias) = 0 Then rawAlias = vendor
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Chon 1 PDF mau de luu tham chieu"
        .Filters.Clear: .Filters.Add "PDF files", "*.pdf"
        If .Show <> -1 Then Exit Sub
        samplePdf = .SelectedItems(1)
    End With
    profileId = "P_" & Replace(NormalizeText(vendor), " ", "_")
    Set vendorMap = ThisWorkbook.Worksheets("NCC_MAP").ListObjects("tblVendorMap")
    Set profiles = ThisWorkbook.Worksheets("PARSER_PROFILES").ListObjects("tblParserProfiles")
    UpsertVendorMapDirect vendorMap, vendor, taxCode, rawAlias, profileId
    UpsertProfileDirect profiles, profileId, vendor, taxCode, samplePdf
    ThisWorkbook.Worksheets("PARSER_PROFILES").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("PARSER_PROFILES").Activate
    MsgBox "Da cau hinh NCC theo MST " & taxCode & ". Tool se map hoa don co MST nay ve " & vendor & " va dung parser mac dinh. PDF mau da luu tai PARSER_PROFILES. Neu parser mac dinh doc sai, dien pattern o 3 cot InvoiceNoPattern, DatePattern, LinePattern; de trong neu parser mac dinh doc dung.", vbInformation
End Sub

Public Sub ThemNCCVaMauHoaDon()
    CauHinhNCCVaParser
End Sub

Private Sub UpsertVendorMapDirect(ByVal lo As ListObject, ByVal vendor As String, ByVal taxCode As String, ByVal rawAlias As String, ByVal profileId As String)
    Dim r As ListRow, rowIndex As Long
    rowIndex = FindRow(lo, 1, vendor)
    If rowIndex = 0 Then Set r = lo.ListRows.Add Else Set r = lo.ListRows(rowIndex)
    r.Range.Cells(1, 1).Resize(1, 6).Value = Array(vendor, taxCode, rawAlias, vendor, profileId, True)
    r.Range.Cells(1, 2).NumberFormat = "@": r.Range.Cells(1, 2).Value = taxCode
End Sub

Private Sub UpsertProfileDirect(ByVal lo As ListObject, ByVal profileId As String, ByVal vendor As String, ByVal taxCode As String, ByVal samplePdf As String)
    Dim r As ListRow, rowIndex As Long
    rowIndex = FindRow(lo, 1, profileId)
    If rowIndex = 0 Then Set r = lo.ListRows.Add Else Set r = lo.ListRows(rowIndex)
    r.Range.Cells(1, 1).Resize(1, 9).Value = Array(profileId, vendor, taxCode, "", "", "", "ACTIVE", samplePdf, "MST la khoa nhan dien NCC. Pattern de trong = dung parser mac dinh; chi dien pattern khi PDF mau doc sai.")
    r.Range.Cells(1, 3).NumberFormat = "@": r.Range.Cells(1, 3).Value = taxCode
End Sub

Private Function VendorAliasExists(ByVal lo As ListObject, ByVal rawVendor As String) As Boolean
    Dim i As Long, normalized As String
    normalized = NormalizeText(rawVendor)
    If lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If NormalizeText(CStr(lo.DataBodyRange.Cells(i, 3).Value)) = normalized Then VendorAliasExists = True: Exit Function
    Next i
End Function

Private Sub AddVendorMapRow(ByVal lo As ListObject, ByVal canonical As String, ByVal rawVendor As String)
    Dim r As ListRow
    Set r = lo.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 6).Value = Array(canonical, "", rawVendor, canonical, "", True)
End Sub

Private Sub AddVendorConfigRow(ByVal lo As ListObject, ByVal canonical As String, ByVal rawVendor As String)
    Dim r As ListRow
    Set r = lo.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 11).Value = Array(canonical, "", rawVendor, canonical, "", "", "", "", "DRAFT", True, "Tu dong tao tu file P. Dien MST neu can map theo MST tren hoa don.")
End Sub

Private Sub UpsertVendorMap(ByVal lo As ListObject, ByVal source As ListObject, ByVal sourceRow As Long)
    Dim rowIndex As Long, r As ListRow
    rowIndex = FindRow(lo, 1, CStr(source.DataBodyRange.Cells(sourceRow, 1).Value))
    If rowIndex = 0 Then Set r = lo.ListRows.Add Else Set r = lo.ListRows(rowIndex)
    r.Range.Cells(1, 1).Resize(1, 6).Value = Array(CStr(source.DataBodyRange.Cells(sourceRow, 1).Value), CStr(source.DataBodyRange.Cells(sourceRow, 2).Value), CStr(source.DataBodyRange.Cells(sourceRow, 3).Value), CStr(source.DataBodyRange.Cells(sourceRow, 4).Value), CStr(source.DataBodyRange.Cells(sourceRow, 5).Value), CBool(source.DataBodyRange.Cells(sourceRow, 10).Value))
    r.Range.Cells(1, 2).NumberFormat = "@"
    r.Range.Cells(1, 2).Value = CStr(source.DataBodyRange.Cells(sourceRow, 2).Value)
End Sub

Private Sub UpsertParserProfile(ByVal lo As ListObject, ByVal source As ListObject, ByVal sourceRow As Long)
    Dim rowIndex As Long, r As ListRow
    rowIndex = FindRow(lo, 1, CStr(source.DataBodyRange.Cells(sourceRow, 5).Value))
    If rowIndex = 0 Then Set r = lo.ListRows.Add Else Set r = lo.ListRows(rowIndex)
    r.Range.Cells(1, 1).Resize(1, 9).Value = Array(CStr(source.DataBodyRange.Cells(sourceRow, 5).Value), CStr(source.DataBodyRange.Cells(sourceRow, 1).Value), CStr(source.DataBodyRange.Cells(sourceRow, 2).Value), CStr(source.DataBodyRange.Cells(sourceRow, 6).Value), CStr(source.DataBodyRange.Cells(sourceRow, 7).Value), CStr(source.DataBodyRange.Cells(sourceRow, 8).Value), CStr(source.DataBodyRange.Cells(sourceRow, 9).Value), "", CStr(source.DataBodyRange.Cells(sourceRow, 11).Value))
    r.Range.Cells(1, 3).NumberFormat = "@"
    r.Range.Cells(1, 3).Value = CStr(source.DataBodyRange.Cells(sourceRow, 2).Value)
End Sub

Private Function FindRow(ByVal lo As ListObject, ByVal columnIndex As Long, ByVal key As String) As Long
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, columnIndex).Value) = key Then FindRow = i: Exit Function
    Next i
End Function
