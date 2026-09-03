Attribute VB_Name = "modDiagnostics"
Option Explicit

Private mDiagnosticExportOk As Boolean

Public Sub ExportDiagnosticBundle()
    Dim outputFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Choose a folder for the VAT Matcher diagnostic bundle"
        If .Show <> -1 Then Exit Sub
        outputFolder = .SelectedItems(1)
    End With
    mDiagnosticExportOk = False
    ExportDiagnosticBundleToFolder outputFolder
    If mDiagnosticExportOk Then
        MsgBox "Diagnostic bundle created. LOG da duoc luu trong goi va xoa khoi workbook de bat dau lich su moi.", vbInformation
    Else
        MsgBox "Diagnostic export failed. Review LOG.", vbCritical
    End If
End Sub

Public Sub ExportDiagnosticBundleToFolder(ByVal outputFolder As String)
    Dim destination As Workbook, sheetNames As Variant, i As Long, outputPath As String, fso As Object
    Dim originalAlerts As Boolean, errorDetail As String
    On Error GoTo EH
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputFolder) Then Err.Raise vbObjectError + 930, , "Diagnostic output folder does not exist."
    outputPath = outputFolder & "\VAT_Matcher_Diagnostic_" & Format$(Now, "yyyymmdd-hhnnss") & ".xlsx"
    If fso.FileExists(outputPath) Then Err.Raise vbObjectError + 931, , "Diagnostic filename already exists. Try again."
    sheetNames = Array("LOG", "BC_HOA_DON", "BC_PHIEU", "VAT_LINES", "PDF_FILES", "EMAIL_HINTS", "MATERIAL_SCOPE_MAP", "CONFIG", "TEST_RESULTS")
    originalAlerts = Application.DisplayAlerts
    'Write this before copying so the bundle contains the whole history up to
    'the moment it is archived. The source LOG is cleared only after SaveAs.
    LogEvent "INFO", "ExportDiagnosticBundleToFolder", "DIAGNOSTIC_EXPORT_ARCHIVING", outputPath, "", "LOG will be cleared only after this workbook is saved successfully."
    Set destination = Workbooks.Add(xlWBATWorksheet)
    For i = LBound(sheetNames) To UBound(sheetNames)
        If i > LBound(sheetNames) Then destination.Worksheets.Add After:=destination.Worksheets(destination.Worksheets.Count)
        CopyDiagnosticSheetValues ThisWorkbook.Worksheets(CStr(sheetNames(i))), destination.Worksheets(destination.Worksheets.Count), CStr(sheetNames(i))
    Next i
    destination.SaveAs outputPath, xlOpenXMLWorkbook
    destination.Close SaveChanges:=True
    ClearDiagnosticLogHistory
    mDiagnosticExportOk = True
    Exit Sub
EH:
    errorDetail = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = originalAlerts
    If Not destination Is Nothing Then destination.Close SaveChanges:=False
    On Error GoTo 0
    LogError "ExportDiagnosticBundleToFolder", "DIAGNOSTIC_EXPORT_FAILED", errorDetail, outputFolder
End Sub

Private Sub ClearDiagnosticLogHistory()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("LOG").ListObjects("tblLog")
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
End Sub

Private Sub CopyDiagnosticSheetValues(ByVal sourceSheet As Worksheet, ByVal destinationSheet As Worksheet, ByVal destinationName As String)
    Dim sourceRange As Range
    Set sourceRange = sourceSheet.UsedRange
    destinationSheet.Name = destinationName
    destinationSheet.Range("A1").Resize(sourceRange.Rows.Count, sourceRange.Columns.Count).Value2 = sourceRange.Value2
    destinationSheet.Cells.Font.Name = "Arial"
    destinationSheet.Columns.AutoFit
End Sub
