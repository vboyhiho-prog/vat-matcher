Attribute VB_Name = "modDashboard"
Option Explicit

Public Sub SetupDashboard()
    Dim ws As Worksheet
    On Error GoTo EH
    Set ws = EnsureDashboardSheet()
    MigrateLegacyPaths
    Application.ScreenUpdating = False
    'Rebuild only the dashboard canvas; never touch technical sheets or tables.
    ws.Range("A1:J40").UnMerge
    ws.Range("A1:J40").Clear
    DeleteDashboardButtons ws
    ws.Cells.Font.Name = "Arial"
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 4
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 4
    ws.Columns("I").ColumnWidth = 18
    ws.Columns("J").ColumnWidth = 18
    ws.Rows("1:40").RowHeight = 20
    ws.Rows("1:2").RowHeight = 24
    ws.Range("A1:J2").Merge
    ws.Range("A1").Value = Vn("0048_1EC6_0020_0054_0048_1ED0_004E_0047_0020_0110_1ED0_0049_0020_0053_004F_00C1_0054_0020_0048_00D3_0041_0020_0110_01A0_004E_0020_0056_0041_0054_0020_0026_0020_0050_0048_0049_1EBE_0055_0020_004E_0048_1EAC_0050_0020_0050")
    ws.Range("A1").Font.Size = 20: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = vbWhite
    ws.Range("A1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A4:J4").Merge
    ws.Range("A4").Value = Vn("004C_00E0_006D_0020_0074_0068_0065_006F_0020_0074_0068_1EE9_0020_0074_1EF1_0020_0063_00E1_0063_0020_006E_00FA_0074_002E_0020_004E_00FA_0074_0020_006E_1EA1_0070_0020_0050_0020_006C_0075_00F4_006E_0020_0063_0068_006F_0020_0070_0068_00E9_0070_0020_0063_0068_1ECD_006E_0020_0066_0069_006C_0065_0020_006D_1EDB_0069_0020_0076_00E0_0020_006E_1EA1_0070_0020_006C_1EA1_0069_0020_0064_1EEF_0020_006C_0069_1EC7_0075_002E")
    ws.Range("A6:J6").Merge: ws.Range("A6").Value = Vn("0031_002E_0020_0043_1EA4_0055_0020_0048_00CC_004E_0048_0020_0044_1EEE_0020_004C_0049_1EC6_0055"): ws.Range("A6").Font.Bold = True: ws.Range("A6").Interior.Color = RGB(221, 235, 247)
    ws.Range("A7:B7").Merge: ws.Range("A7").Value = Vn("0054_0068_01B0_0020_006D_1EE5_0063_0020_0065_006D_0061_0069_006C")
    ws.Range("C7:J7").Merge
    ws.Range("A8:B8").Merge: ws.Range("A8").Value = Vn("0054_0068_01B0_0020_006D_1EE5_0063_0020_0050_0044_0046")
    ws.Range("C8:J8").Merge
    ws.Range("A9:B9").Merge: ws.Range("A9").Value = Vn("0046_0069_006C_0065_0020_0054_0068_0065_006F_0020_0064_00F5_0069_0020_0050")
    ws.Range("C9:J9").Merge
    ws.Range("A10:B10").Merge: ws.Range("A10").Value = Vn("0054_0072_1EA1_006E_0067_0020_0074_0068_00E1_0069_0020_006E_1EA1_0070_0020_0050")
    ws.Range("C10:J10").Merge: ws.Range("C10").Value = Vn("0043_0048_01AF_0041_0020_004E_1EA0_0050_0020_0046_0049_004C_0045_0020_0050")
    ws.Range("A7:B9").Font.Bold = True
    ws.Range("A7:B10").Font.Bold = True
    ws.Range("A11:C11").Merge: ws.Range("A11").Value = Vn("0032_002E_0020_0054_0048_1ED0_004E_0047_0020_004B_00CA_0020_004B_1EBE_0054_0020_0051_0055_1EA2_0020_0047_1EA6_004E_0020_004E_0048_1EA4_0054"): ws.Range("A11").Font.Bold = True: ws.Range("A11").Interior.Color = RGB(221, 235, 247)
    ws.Range("A12:B12").Merge: ws.Range("A12").Value = Vn("0050_0044_0046_0020_0111_00E3_0020_0074_0072_00ED_0063_0068_0020_0074_1EEB_0020_0065_006D_0061_0069_006C"): ws.Range("C12").Value = 0
    ws.Range("A13:B13").Merge: ws.Range("A13").Value = Vn("0050_0044_0046_0020_0111_00E3_0020_0071_0075_00E9_0074"): ws.Range("C13").Value = 0
    ws.Range("A14:B14").Merge: ws.Range("A14").Value = Vn("004B_0068_1EDB_0070_0020_0068_006F_00E0_006E_0020_0074_006F_00E0_006E"): ws.Range("C14").Value = 0
    ws.Range("A15:B15").Merge: ws.Range("A15").Value = Vn("004C_1EAB_006E_0020_002F_0020_006E_0067_006F_00E0_0069_0020_0078_01B0_1EDF_006E_0067"): ws.Range("C15").Value = 0
    ws.Range("A16:B16").Merge: ws.Range("A16").Value = Vn("0043_1EA7_006E_0020_0078_00E1_0063_0020_006D_0069_006E_0068"): ws.Range("C16").Value = 0
    ws.Range("A12:B16").Font.Bold = True: ws.Range("C12:C16").Font.Size = 14: ws.Range("C12:C16").Font.Bold = True
    ws.Range("E11:G11").Merge: ws.Range("E11").Value = Vn("0042_00C1_004F_0020_0043_00C1_004F_0020_0026_0020_0043_1EA4_0055_0020_0048_00CC_004E_0048"): ws.Range("E11").Font.Bold = True: ws.Range("E11").Interior.Color = RGB(221, 235, 247)
    ws.Range("E12:F12").Merge: ws.Range("E12").Value = Vn("0042_00E1_006F_0020_0063_00E1_006F_0020_0068_00F3_0061_0020_0111_01A1_006E"): ws.Hyperlinks.Add Anchor:=ws.Range("G12"), Address:="", SubAddress:="'BC_HOA_DON'!A1", TextToDisplay:=Vn("004D_1EDE")
    ws.Range("E13:F13").Merge: ws.Range("E13").Value = Vn("0042_00E1_006F_0020_0063_00E1_006F_0020_0070_0068_0069_1EBF_0075"): ws.Hyperlinks.Add Anchor:=ws.Range("G13"), Address:="", SubAddress:="'BC_PHIEU'!A1", TextToDisplay:=Vn("004D_1EDE")
    ws.Range("E14:F14").Merge: ws.Range("E14").Value = Vn("004E_0043_0043_0020_002F_0020_006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0"): ws.Hyperlinks.Add Anchor:=ws.Range("G14"), Address:="", SubAddress:="'NCC_CONFIG_IMPORT'!A1", TextToDisplay:=Vn("004D_1EDE")
    ws.Range("E15:F15").Merge: ws.Range("E15").Value = Vn("004E_0068_1EAD_0074_0020_006B_00FD_0020_0068_006F_1EA1_0074_0020_0111_1ED9_006E_0067"): ws.Hyperlinks.Add Anchor:=ws.Range("G15"), Address:="", SubAddress:="'LOG'!A1", TextToDisplay:=Vn("004D_1EDE")
    ws.Range("A19:J19").Merge: ws.Range("A19").Value = Vn("0033_002E_0020_0042_1EA2_004E_0047_0020_0110_0049_1EC0_0055_0020_004B_0048_0049_1EC2_004E"): ws.Range("A19").Font.Bold = True: ws.Range("A19").Interior.Color = RGB(221, 235, 247)
    AddDashboardButton ws, Vn("0031_002E_0020_0043_0048_1ECC_004E_0020_0045_004D_0041_0049_004C_0020_002B_0020_0054_0052_00CD_0043_0048_0020_0050_0044_0046"), "DashboardExtractEmailFolder", 20, 385, 250
    AddDashboardButton ws, Vn("0032_002E_0020_0043_0048_1ECC_004E_0020_0054_0048_01AF_0020_004D_1EE4_0043_0020_0050_0044_0046"), "DashboardChoosePdfFolder", 300, 385, 250
    AddDashboardButton ws, Vn("0033_002E_0020_0043_0048_1ECC_004E_0020_0026_0020_004E_1EA0_0050_0020_004C_1EA0_0049_0020_0046_0049_004C_0045_0020_0050"), "DashboardLoadTrackingOnce", 580, 385, 250
    AddDashboardButton ws, Vn("0034_002E_0020_004B_0049_1EC2_004D_0020_0054_0052_0041_0020_0026_0020_004B_0048_1EDA_0050_0020_0050_0048_0049_1EBE_0055"), "DashboardRunMatching", 20, 430, 250
    AddDashboardButton ws, Vn("0035_002E_0020_0058_0045_004D_0020_0054_0052_01AF_1EDA_0043_0020_002F_0020_0110_1ED4_0049_0020_0054_00CA_004E"), "DashboardRename", 300, 430, 250
    AddDashboardButton ws, Vn("0043_1EAC_0050_0020_004E_0048_1EAC_0054_0020_0054_0048_1ED0_004E_0047_0020_004B_00CA"), "UpdateDashboard", 580, 430, 250
    AddDashboardButton ws, Vn("1EA8_004E_0020_0053_0048_0045_0045_0054_0020_004B_1EF8_0020_0054_0048_0055_1EAC_0054"), "HideTechnicalSheets", 20, 475, 250
    AddDashboardButton ws, Vn("0058_0055_1EA4_0054_0020_0047_00D3_0049_0020_0043_0048_1EA8_004E_0020_0110_004F_00C1_004E"), "ExportDiagnosticBundle", 300, 475, 250
    AddDashboardButton ws, "6. DOI TEN DA DUYET", "DashboardApplyRenames", 580, 475, 250
    AddDashboardButton ws, "XOA DU LIEU PHIEN CU", "ClearOldSessionData", 20, 520, 250
    AddDashboardButton ws, "CAU HINH NCC + PARSER", "CauHinhNCCVaParser", 300, 520, 250
    UpdateDashboard
    ws.Activate
    Application.ScreenUpdating = True
    Exit Sub
EH:
    LogError "SetupDashboard", "DASHBOARD_SETUP_FAILED", Err.Description
    Application.ScreenUpdating = True
End Sub

Public Sub DashboardExtractEmailFolder()
    Dim folderPath As String, fso As Object, folder As Object, file As Object, destination As String, extractedBefore As Long, extractedAfter As Long
    On Error GoTo EH
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder containing email .msg files"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    extractedBefore = CountTableStatus("EMAIL_ATTACHMENTS", "tblEmailAttachments", 5, "EXTRACTED")
    Set fso = CreateObject("Scripting.FileSystemObject"): Set folder = fso.GetFolder(folderPath)
    destination = folderPath & "\PDF_Extracted"
    If Not fso.FolderExists(destination) Then fso.CreateFolder destination
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "msg" Then ExtractPdfAttachmentsFromMsgPath CStr(file.Path), destination
    Next file
    DashboardSheet.Range("C7").Value = folderPath
    UpdateDashboard
    extractedAfter = CountTableStatus("EMAIL_ATTACHMENTS", "tblEmailAttachments", 5, "EXTRACTED")
    MsgBox "Da tai " & CStr(extractedAfter - extractedBefore) & " file PDF tu email. Tong PDF da tai: " & CStr(extractedAfter) & ".", vbInformation
    Exit Sub
EH:
    LogError "DashboardExtractEmailFolder", "DASHBOARD_EMAIL_EXTRACT_FAILED", Err.Description, folderPath
    MsgBox "Email PDF extraction failed. Review LOG.", vbCritical
End Sub

Public Sub DashboardChoosePdfFolder()
    ConfigurePdfFolderBatch
    DashboardSheet.Range("C8").Value = PdfFolderPath()
    UpdateDashboard
End Sub

Public Sub DashboardLoadTrackingOnce()
    Dim selectedFile As String, priorPath As String, lo As ListObject
    On Error GoTo EH
    priorPath = TrackingPPath()
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Choose the latest Tracking P file to load or reload"
        .Filters.Clear: .Filters.Add "Excel files", "*.xlsx;*.xlsm;*.xls"
        If Len(priorPath) > 0 Then .InitialFileName = priorPath
        If .Show <> -1 Then Exit Sub
        selectedFile = .SelectedItems(1)
    End With
    DashboardSheet.Range("C10").Value = "LOADING: " & selectedFile
    DashboardSheet.Range("C10").Interior.Color = RGB(255, 242, 204)
    DoEvents
    LoadTrackingData selectedFile
    Set lo = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    If lo.ListRows.Count = 0 Or TrackingPPath() <> selectedFile Then Err.Raise vbObjectError + 959, , "The selected Tracking P file could not be loaded."
    UpdateDashboard
    MsgBox "Reloaded " & Format$(lo.ListRows.Count, "#,##0") & " rows from the selected Tracking P file.", vbInformation
    Exit Sub
EH:
    DashboardSheet.Range("C10").Value = "TRACKING P LOAD FAILED - see LOG"
    DashboardSheet.Range("C10").Interior.Color = RGB(244, 204, 204)
    LogError "DashboardLoadTrackingOnce", "DASHBOARD_LOAD_P_FAILED", Err.Description, selectedFile
    MsgBox "Tracking P could not be loaded. Please check LOG.", vbCritical
End Sub

Public Sub DashboardRunMatching()
    Dim pdfCount As Long, matchedCount As Long, mixedCount As Long, reviewCount As Long
    On Error GoTo EH
    If Not TrackingLoadedThisSession() Then Err.Raise vbObjectError + 960, , "Hay chon va nap lai file P trong phien Excel hien tai."
    If Not PdfFolderSelectedThisSession() Then Err.Raise vbObjectError + 961, , "Hay chon lai thu muc PDF trong phien Excel hien tai."
    If ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData").ListRows.Count = 0 Then Err.Raise vbObjectError + 962, , "File P khong co du lieu."
    RunPythonPipeline
    UpdateDashboard
    pdfCount = ThisWorkbook.Worksheets("PDF_FILES").ListObjects("tblPdfFiles").ListRows.Count
    CountPdfReportStates matchedCount, mixedCount, reviewCount
    ShowSheet "BC_HOA_DON"
    MsgBox "Da kiem tra " & CStr(pdfCount) & " file PDF: " & CStr(matchedCount) & " khop, " & CStr(mixedCount) & " lan/ngoai xuong, " & CStr(reviewCount) & " can xac minh.", vbInformation
    Exit Sub
EH:
    LogError "DashboardRunMatching", "DASHBOARD_MATCH_FAILED", Err.Description
    MsgBox "Kiem tra khop that bai. Xem LOG.", vbCritical
End Sub

Public Sub DashboardRename()
    CreateRenamePreviews
    ShowSheet "BC_HOA_DON"
    MsgBox "Da tao ten de xuat. Kiem tra BAO CAO HOA DON, nhap OK o cot Quyet dinh, sau do bam nut DOI TEN DA DUYET. Nhap NG de khong doi ten.", vbInformation
End Sub

Public Sub DashboardApplyRenames()
    ApplyApprovedRenames
End Sub

Public Sub UpdateDashboard()
    Dim ws As Worksheet, emailCount As Long, pdfCount As Long, matchedCount As Long, mixedCount As Long, reviewCount As Long
    On Error GoTo EH
    Set ws = DashboardSheet
    emailCount = CountTableStatus("EMAIL_ATTACHMENTS", "tblEmailAttachments", 5, "EXTRACTED")
    pdfCount = ThisWorkbook.Worksheets("PDF_FILES").ListObjects("tblPdfFiles").ListRows.Count
    CountPdfReportStates matchedCount, mixedCount, reviewCount
    ws.Range("C12").Value = emailCount: ws.Range("C13").Value = pdfCount: ws.Range("C14").Value = matchedCount: ws.Range("C15").Value = mixedCount: ws.Range("C16").Value = reviewCount
    ws.Range("C8").Value = PdfFolderPath()
    ws.Range("C9").Value = TrackingPPath()
    UpdateTrackingLoadStatus ws
    Exit Sub
EH:
    LogError "UpdateDashboard", "DASHBOARD_METRICS_FAILED", Err.Description
End Sub

Private Sub UpdateTrackingLoadStatus(ByVal ws As Worksheet)
    Dim lo As ListObject, filePath As String, rowCount As Long
    Set lo = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    rowCount = lo.ListRows.Count
    filePath = TrackingPPath()
    If rowCount = 0 Or Len(filePath) = 0 Then
        ws.Range("C10").Value = "TRACKING P NOT LOADED"
        ws.Range("C10").Interior.Color = RGB(255, 242, 204)
    Else
        ws.Range("C10").Value = "LOADED " & Format$(rowCount, "#,##0") & " rows at " & Format$(Now, "hh:nn dd/mm/yyyy") & ". Use button 3 to select and reload the latest Tracking P file."
        ws.Range("C10").Interior.Color = RGB(226, 239, 218)
    End If
End Sub

Private Function PdfFolderPath() As String
    PdfFolderPath = CStr(ThisWorkbook.Worksheets("HOME").Range("B4").Value)
End Function

Private Function TrackingPPath() As String
    TrackingPPath = CStr(ThisWorkbook.Worksheets("HOME").Range("D4").Value)
End Function

Private Sub MigrateLegacyPaths()
    Dim home As Worksheet, legacyPath As String, legacyTracking As String, pdfs As ListObject, fullPath As String
    Set home = ThisWorkbook.Worksheets("HOME")
    If Len(CStr(home.Range("B4").Value)) > 0 Or Len(CStr(home.Range("D4").Value)) > 0 Then Exit Sub
    legacyPath = CStr(home.Range("B2").Value)
    legacyTracking = CStr(home.Range("D3").Value)
    If LCase$(Right$(legacyPath, 5)) = ".xlsx" Or LCase$(Right$(legacyPath, 5)) = ".xlsm" Or LCase$(Right$(legacyPath, 4)) = ".xls" Then
        home.Range("D4").Value = legacyPath
        Set pdfs = ThisWorkbook.Worksheets("PDF_FILES").ListObjects("tblPdfFiles")
        If Not pdfs.DataBodyRange Is Nothing Then fullPath = CStr(pdfs.DataBodyRange.Cells(1, 2).Value)
        If Len(fullPath) > 0 Then home.Range("B4").Value = Left$(fullPath, InStrRev(fullPath, "\") - 1)
    Else
        If InStr(1, legacyPath, "\", vbTextCompare) > 0 Then home.Range("B4").Value = legacyPath
        If Len(legacyTracking) > 0 Then home.Range("D4").Value = legacyTracking
    End If
End Sub

Public Sub HideTechnicalSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "DASHBOARD" And ws.Name <> "BC_HOA_DON" And ws.Name <> "BC_PHIEU" And ws.Name <> "NCC_CONFIG_IMPORT" And ws.Name <> "MATERIAL_SCOPE_MAP" And ws.Name <> "LOG" Then ws.Visible = xlSheetHidden
    Next ws
    DashboardSheet.Activate
End Sub

Public Sub ShowSheet(ByVal sheetName As String)
    ThisWorkbook.Worksheets(sheetName).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(sheetName).Activate
End Sub

Private Function DashboardSheet() As Worksheet
    Set DashboardSheet = EnsureDashboardSheet()
End Function

'VBIDE imports .bas files through the local ANSI code page.  Keep Vietnamese UI text as
'Unicode code points so the dashboard renders correctly on every corporate Windows locale.
Private Function Vn(ByVal codePoints As String) As String
    Dim parts() As String, i As Long
    parts = Split(codePoints, "_")
    For i = LBound(parts) To UBound(parts)
        Vn = Vn & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function

Private Function EnsureDashboardSheet() As Worksheet
    On Error Resume Next: Set EnsureDashboardSheet = ThisWorkbook.Worksheets("DASHBOARD"): On Error GoTo 0
    If EnsureDashboardSheet Is Nothing Then Set EnsureDashboardSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1)): EnsureDashboardSheet.Name = "DASHBOARD"
End Function

Private Sub AddDashboardButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal leftPos As Double, ByVal topPos As Double, ByVal width As Double)
    Dim button As Shape
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, width, 30)
    button.Name = "dash_" & Replace(macroName, " ", "_")
    'Use the legacy TextFrame API: it is supported consistently by corporate Excel builds.
    button.TextFrame.Characters.Text = caption
    button.TextFrame.Characters.Font.Size = 10
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.HorizontalAlignment = xlHAlignCenter
    button.TextFrame.VerticalAlignment = xlVAlignCenter
    button.Fill.ForeColor.RGB = RGB(31, 78, 121): button.Line.ForeColor.RGB = RGB(31, 78, 121)
    button.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    button.OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
End Sub

Private Sub DeleteDashboardButtons(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "dash_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Function CountTableStatus(ByVal sheetName As String, ByVal tableName As String, ByVal statusColumn As Long, ByVal wantedStatus As String) As Long
    Dim lo As ListObject, i As Long
    Set lo = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If UCase$(CStr(lo.DataBodyRange.Cells(i, statusColumn).Value)) = wantedStatus Then CountTableStatus = CountTableStatus + 1
    Next i
End Function

Private Sub CountPdfReportStates(ByRef matchedCount As Long, ByRef mixedCount As Long, ByRef reviewCount As Long)
    Dim lo As ListObject, i As Long, pdf As Variant, status As String, reason As String, allMatched As Object, mixed As Object, review As Object
    Set lo = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set allMatched = CreateObject("Scripting.Dictionary"): Set mixed = CreateObject("Scripting.Dictionary"): Set review = CreateObject("Scripting.Dictionary")
    For i = 1 To lo.ListRows.Count
        pdf = CStr(lo.DataBodyRange.Cells(i, 2).Value): status = CStr(lo.DataBodyRange.Cells(i, 8).Value): reason = CStr(lo.DataBodyRange.Cells(i, 11).Value)
        If Not allMatched.Exists(pdf) Then allMatched.Add pdf, True
        If status <> "MATCHED" Then allMatched(pdf) = False
        If status = "OTHER_FACTORY" Or InStr(1, reason, "MIXED_SCOPE", vbTextCompare) > 0 Then mixed(pdf) = True
        If status <> "MATCHED" And status <> "OTHER_FACTORY" Then review(pdf) = True
    Next i
    For Each pdf In allMatched.Keys
        If mixed.Exists(pdf) Then
            mixedCount = mixedCount + 1
        ElseIf allMatched(pdf) Then
            matchedCount = matchedCount + 1
        End If
    Next pdf
    reviewCount = review.Count
End Sub
