Attribute VB_Name = "modRenameRollback"
Option Explicit

Public Sub ApplyApprovedRenames()
    Dim workFolder As String, renamedCount As Long, okRows As Long, skippedRows As Long
    CreateRenamePreviews
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Chon thu muc BAN SAO chua PDF can doi ten"
        If .Show <> -1 Then Exit Sub
        workFolder = .SelectedItems(1)
    End With
    renamedCount = ApplyApprovedRenamesInFolder(workFolder, okRows, skippedRows)
    MsgBox "Da doi ten " & CStr(renamedCount) & " PDF tu " & CStr(okRows) & " dong OK. Bo qua " & CStr(skippedRows) & " dong (trung PDF, thieu file hoac ten da ton tai). NG se khong doi.", vbInformation
End Sub

Public Function ApplyApprovedRenamesInFolder(ByVal workFolder As String, ByRef okRows As Long, ByRef skippedRows As Long) As Long
    Dim report As ListObject, i As Long, decision As String, preview As String, originalName As String, renamedCount As Long, renamedSources As Object
    Dim originalPath As String, newPath As String
    On Error GoTo EH
    EnsureSafeWorkFolder workFolder
    Set report = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    If report.DataBodyRange Is Nothing Then Exit Function
    Set renamedSources = CreateObject("Scripting.Dictionary")
    For i = 1 To report.ListRows.Count
        decision = UCase$(CStr(report.DataBodyRange.Cells(i, 10).Value))
        preview = CStr(report.DataBodyRange.Cells(i, 9).Value)
        If decision = "OK" Then
            okRows = okRows + 1
            If Len(preview) = 0 Then
                skippedRows = skippedRows + 1
                LogEvent "WARN", "ApplyApprovedRenamesInFolder", "RENAME_PREVIEW_MISSING", CStr(report.DataBodyRange.Cells(i, 3).Value), "", "Tao ten de xuat truoc khi doi ten."
            Else
                originalName = OriginalNameForInvoice(CStr(report.DataBodyRange.Cells(i, 3).Value))
                originalPath = workFolder & "\" & originalName
                newPath = workFolder & "\" & preview
                If renamedSources.Exists(LCase$(originalPath)) Then
                    'One PDF may contain several invoices; all rows use the same proposed name and are renamed once.
                    skippedRows = skippedRows + 1
                ElseIf Len(Dir$(originalPath)) = 0 Then
                    skippedRows = skippedRows + 1
                    LogEvent "WARN", "ApplyApprovedRenamesInFolder", "RENAME_SOURCE_MISSING", originalName, originalPath, newPath
                ElseIf Len(Dir$(newPath)) > 0 Then
                    skippedRows = skippedRows + 1
                    LogEvent "WARN", "ApplyApprovedRenamesInFolder", "RENAME_TARGET_EXISTS", preview, originalPath, newPath
                Else
                    Name originalPath As newPath
                    renamedCount = renamedCount + 1
                    renamedSources.Add LCase$(originalPath), True
                    LogEvent "INFO", "ApplyApprovedRenamesInFolder", "RENAME_APPLIED", CStr(report.DataBodyRange.Cells(i, 3).Value) & " approved by " & Environ$("Username"), originalPath, newPath
                End If
            End If
        End If
    Next i
    ApplyApprovedRenamesInFolder = renamedCount
    Exit Function
EH:
    LogError "ApplyApprovedRenamesInFolder", "RENAME_FAILED", Err.Description, workFolder
    MsgBox "Doi ten that bai. Xem LOG.", vbCritical
End Function

Public Sub RollbackRenames()
    Dim workFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the same copied work folder used for rename"
        If .Show <> -1 Then Exit Sub
        workFolder = .SelectedItems(1)
    End With
    RollbackRenamesInFolder workFolder
    MsgBox "Rename rollback completed. Review LOG.", vbInformation
End Sub

Public Sub RollbackRenamesInFolder(ByVal workFolder As String)
    Dim lo As ListObject, i As Long, originalPath As String, newPath As String
    On Error GoTo EH
    EnsureSafeWorkFolder workFolder
    Set lo = ThisWorkbook.Worksheets("LOG").ListObjects("tblLog")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = lo.ListRows.Count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(i, 6).Value) = "RENAME_APPLIED" Then
            originalPath = CStr(lo.DataBodyRange.Cells(i, 8).Value)
            newPath = CStr(lo.DataBodyRange.Cells(i, 9).Value)
            If IsPathInFolder(originalPath, workFolder) And IsPathInFolder(newPath, workFolder) Then
                If Len(Dir$(newPath)) > 0 And Len(Dir$(originalPath)) = 0 Then
                    Name newPath As originalPath
                    LogEvent "INFO", "RollbackRenamesInFolder", "RENAME_ROLLED_BACK", "Rollback completed.", newPath, originalPath
                End If
            End If
        End If
    Next i
    Exit Sub
EH:
    LogError "RollbackRenamesInFolder", "ROLLBACK_FAILED", Err.Description, workFolder
    MsgBox "Rollback failed. Review LOG.", vbCritical
End Sub

Private Function IsPathInFolder(ByVal fullPath As String, ByVal folderPath As String) As Boolean
    IsPathInFolder = LCase$(Left$(fullPath, Len(folderPath) + 1)) = LCase$(folderPath & "\")
End Function

Private Sub EnsureSafeWorkFolder(ByVal workFolder As String)
    Dim fso As Object, normalizedFolder As String
    normalizedFolder = Trim$(workFolder)
    If Len(normalizedFolder) = 0 Then Err.Raise vbObjectError + 890, , "Chua chon thu muc PDF can doi ten."
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(normalizedFolder) Then Err.Raise vbObjectError + 891, , "Thu muc PDF khong ton tai: " & normalizedFolder
End Sub

Private Function OriginalNameForInvoice(ByVal invoiceNo As String) As String
    Dim invoices As ListObject, pdfs As ListObject, i As Long, pdfId As String, fullPath As String
    Set invoices = ThisWorkbook.Worksheets("INVOICES").ListObjects("tblInvoices")
    Set pdfs = ThisWorkbook.Worksheets("PDF_FILES").ListObjects("tblPdfFiles")
    For i = 1 To invoices.ListRows.Count
        If CStr(invoices.DataBodyRange.Cells(i, 5).Value) = invoiceNo Then pdfId = CStr(invoices.DataBodyRange.Cells(i, 2).Value): Exit For
    Next i
    For i = 1 To pdfs.ListRows.Count
        If CStr(pdfs.DataBodyRange.Cells(i, 1).Value) = pdfId Then fullPath = CStr(pdfs.DataBodyRange.Cells(i, 2).Value): Exit For
    Next i
    OriginalNameForInvoice = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
End Function
