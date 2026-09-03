Attribute VB_Name = "modPdfBatch"
Option Explicit

' The selected folder is intentionally only a path in the Excel UI. PDF
' extraction is owned by the bundled portable engine, so no Office PDF connector is used.
Private mPdfFolderSelectedThisSession As Boolean

Public Function PdfFolderSelectedThisSession() As Boolean
    PdfFolderSelectedThisSession = mPdfFolderSelectedThisSession
End Function

Public Sub ConfigurePdfFolderBatch()
    Dim folderPath As String
    On Error GoTo EH
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Chon thu muc chua PDF hoa don VAT"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    ConfigurePdfFolderBatchPath folderPath
    MsgBox "Da chon thu muc PDF. Nut KIEM TRA & KHOP PHIEU se goi Python doc PDF.", vbInformation
    Exit Sub
EH:
    LogError "ConfigurePdfFolderBatch", "PDF_FOLDER_SETUP_FAILED", Err.Description, folderPath
    MsgBox "Khong the chon thu muc PDF. Xem LOG.", vbCritical
End Sub

Public Sub ConfigurePdfFolderBatchPath(ByVal folderPath As String)
    On Error GoTo EH
    mPdfFolderSelectedThisSession = False
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Err.Raise vbObjectError + 730, , "Thu muc PDF khong ton tai."
    ThisWorkbook.Worksheets("HOME").Range("B4").Value = folderPath
    mPdfFolderSelectedThisSession = True
    LogEvent "INFO", "ConfigurePdfFolderBatchPath", "PYTHON_PDF_FOLDER_READY", "Thu muc PDF da san sang cho Python parser.", folderPath, "Chon va nap file P, sau do chay doi soat."
    Exit Sub
EH:
    LogError "ConfigurePdfFolderBatchPath", "PDF_FOLDER_SETUP_FAILED", Err.Description, folderPath
    Err.Raise Err.Number, , Err.Description
End Sub

Public Sub RefreshPdfBatch()
    'Compatibility entry point for old macros. It must not invoke Power Query.
    LogEvent "INFO", "RefreshPdfBatch", "PDF_PIPELINE_PYTHON", "Power Query PDF da duoc thay bang Python parser.", PdfFolderPath(), "Chay doi soat tu Dashboard."
End Sub

Public Sub ResetPdfQueryForRelease()
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PQ_PDF_RAW")
    ws.ListObjects("tblPdfRaw").Delete
    ThisWorkbook.Connections("Query - VAT_PDF_BATCH").Delete
    ThisWorkbook.Queries("VAT_PDF_BATCH").Delete
    On Error GoTo 0
    ws.Range("A2:I2").Value = Array("Name", "Extension", "Folder Path", "Date modified", "PdfTableId", "PdfTableKind", "PdfTableName", "PdfTableData", "PdfText")
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A2:I2"), , xlYes)
    lo.Name = "tblPdfRaw"
    lo.TableStyle = "TableStyleMedium2"
End Sub

Public Function PdfFolderPath() As String
    PdfFolderPath = Trim$(CStr(ThisWorkbook.Worksheets("HOME").Range("B4").Value))
End Function
