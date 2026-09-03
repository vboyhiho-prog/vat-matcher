Attribute VB_Name = "modEmailIntake"
Option Explicit

Public Sub ExtractPdfAttachmentsFromMsg()
    Dim msgPath As String, baseFolder As String, destinationFolder As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select an Outlook MSG file"
        .Filters.Clear
        .Filters.Add "Outlook message", "*.msg"
        If .Show <> -1 Then Exit Sub
        msgPath = .SelectedItems(1)
    End With
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select an ASCII output folder for extracted PDFs"
        If .Show <> -1 Then Exit Sub
        baseFolder = .SelectedItems(1)
    End With
    destinationFolder = baseFolder & "\VAT_Email_PDF_" & Format$(Now, "yyyymmdd-hhnnss")
    ExtractPdfAttachmentsFromMsgPath msgPath, destinationFolder
    MsgBox "PDF attachments extracted. Review EMAIL_ATTACHMENTS and LOG.", vbInformation
End Sub

Public Sub ExtractPdfAttachmentsFromMsgPath(ByVal msgPath As String, ByVal destinationFolder As String)
    Dim outlookApp As Object, mailItem As Object, attachment As Object
    Dim lo As ListObject, r As ListRow, i As Long, name As String, targetPath As String, pdfCount As Long
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo EH
    If Len(Dir$(msgPath)) = 0 Then Err.Raise vbObjectError + 850, , "MSG file does not exist."
    If Len(Dir$(destinationFolder, vbDirectory)) = 0 Then MkDir destinationFolder
    Set lo = ThisWorkbook.Worksheets("EMAIL_ATTACHMENTS").ListObjects("tblEmailAttachments")
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.Session.OpenSharedItem(msgPath)
    For i = 1 To mailItem.Attachments.Count
        Set attachment = mailItem.Attachments.Item(i)
        name = CStr(attachment.FileName)
        If LCase$(FileExtension(name)) = "pdf" Then
            targetPath = destinationFolder & "\" & SafeFileName(name)
            Set r = lo.ListRows.Add
            r.Range.Cells(1, 1).Resize(1, 6).Value = Array(msgPath, CStr(mailItem.Subject), name, FileExtension(name), "", targetPath)
            If Len(Dir$(targetPath)) > 0 Then
                If ExistingAttachmentRow(lo, msgPath, name, targetPath) > 0 Then
                    r.Range.Cells(1, 5).Value = "SKIPPED_DUPLICATE"
                    LogEvent "INFO", "ExtractPdfAttachmentsFromMsgPath", "MSG_PDF_SKIPPED_DUPLICATE", name, msgPath, targetPath
                Else
                    r.Range.Cells(1, 5).Value = "NAME_CONFLICT"
                    LogEvent "WARNING", "ExtractPdfAttachmentsFromMsgPath", "MSG_PDF_NAME_CONFLICT", name, msgPath, targetPath
                End If
            Else
                attachment.SaveAsFile targetPath
                r.Range.Cells(1, 5).Value = "EXTRACTED"
                pdfCount = pdfCount + 1
                LogEvent "INFO", "ExtractPdfAttachmentsFromMsgPath", "MSG_PDF_EXTRACTED", name, msgPath, targetPath
            End If
        Else
            Set r = lo.ListRows.Add
            r.Range.Cells(1, 1).Resize(1, 6).Value = Array(msgPath, CStr(mailItem.Subject), name, FileExtension(name), "SKIPPED_NON_PDF", "")
        End If
    Next i
    LogEvent "INFO", "ExtractPdfAttachmentsFromMsgPath", "MSG_EXTRACTION_OK", CStr(pdfCount) & " PDF attachments extracted.", msgPath, destinationFolder
    Exit Sub
EH:
    failureNumber = Err.Number
    failureDescription = Err.Description
    LogError "ExtractPdfAttachmentsFromMsgPath", "MSG_EXTRACTION_FAILED", failureDescription, msgPath
    Err.Raise failureNumber, , failureDescription
End Sub

Private Function ExistingAttachmentRow(ByVal lo As ListObject, ByVal msgPath As String, ByVal name As String, ByVal savedPath As String) As Long
    Dim i As Long
    For i = 1 To lo.ListRows.Count - 1
        If CStr(lo.DataBodyRange.Cells(i, 1).Value) = msgPath And CStr(lo.DataBodyRange.Cells(i, 3).Value) = name And CStr(lo.DataBodyRange.Cells(i, 5).Value) = "EXTRACTED" And CStr(lo.DataBodyRange.Cells(i, 6).Value) = savedPath Then ExistingAttachmentRow = i: Exit Function
    Next i
End Function

Private Function FileExtension(ByVal fileName As String) As String
    Dim position As Long
    position = InStrRev(fileName, ".")
    If position > 0 Then FileExtension = Mid$(fileName, position + 1)
End Function

Private Function SafeFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant, i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SafeFileName = fileName
    For i = LBound(invalidChars) To UBound(invalidChars)
        SafeFileName = Replace(SafeFileName, CStr(invalidChars(i)), "_")
    Next i
End Function
