Attribute VB_Name = "modEmailHints"
Option Explicit

Public Sub ReadMsgHintsFromPath(ByVal msgPath As String)
    Dim outlookApp As Object, mailItem As Object, re As Object, matches As Object, m As Object
    Dim hints As ListObject, gr As ListObject, i As Long, ib As String, contextText As String
    On Error GoTo EH
    If Len(Dir$(msgPath)) = 0 Then Err.Raise vbObjectError + 930, , "MSG file does not exist."
    Set hints = ThisWorkbook.Worksheets("EMAIL_HINTS").ListObjects("tblEmailHints")
    Set gr = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.Session.OpenSharedItem(msgPath)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\b[0-9]{9}\b"
    Set matches = re.Execute(CStr(mailItem.Body))
    For Each m In matches
        ib = CStr(m.Value)
        i = FindGrRowByIb(gr, ib)
        If i > 0 Then
            contextText = Left$(Replace(Replace(CStr(mailItem.Body), vbCr, " "), vbLf, " "), 160)
            AddHighConfidenceHint hints, msgPath, ib, CStr(gr.DataBodyRange.Cells(i, 3).Value), CDbl(Val(CStr(gr.DataBodyRange.Cells(i, 7).Value))), gr.DataBodyRange.Cells(i, 9).Value, CStr(gr.DataBodyRange.Cells(i, 4).Value), contextText
        End If
    Next m
    LogEvent "INFO", "ReadMsgHintsFromPath", "EMAIL_HINTS_OK", CStr(matches.Count) & " IB tokens checked.", msgPath, "Only IB values present in GR_DATA become HIGH hints."
    Exit Sub
EH:
    LogError "ReadMsgHintsFromPath", "EMAIL_HINTS_FAILED", Err.Description, msgPath
End Sub

Private Function FindGrRowByIb(ByVal gr As ListObject, ByVal ib As String) As Long
    Dim i As Long
    For i = 1 To gr.ListRows.Count
        If CStr(gr.DataBodyRange.Cells(i, 8).Value) = ib Then FindGrRowByIb = i: Exit Function
    Next i
End Function

Public Function IbRangeMatchesReceiptForTest(ByVal firstIb As Long, ByVal lastIb As Long, ByVal receiptNo As String) As Boolean
    Dim gr As ListObject, data As Variant, ibIndex As Object, ibValue As Long, i As Long, rowIndex As Long
    Set gr = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    If gr.DataBodyRange Is Nothing Then Exit Function
    data = gr.DataBodyRange.Value2
    Set ibIndex = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 8))) > 0 And Not ibIndex.Exists(CStr(data(i, 8))) Then ibIndex.Add CStr(data(i, 8)), i
    Next i
    For ibValue = firstIb To lastIb
        If Not ibIndex.Exists(CStr(ibValue)) Then Exit Function
        rowIndex = CLng(ibIndex(CStr(ibValue)))
        If CStr(data(rowIndex, 2)) <> receiptNo Then Exit Function
    Next ibValue
    IbRangeMatchesReceiptForTest = True
End Function

Private Sub AddHighConfidenceHint(ByVal hints As ListObject, ByVal msgPath As String, ByVal ib As String, ByVal material As String, ByVal qty As Double, ByVal grDate As Variant, ByVal vendor As String, ByVal contextText As String)
    Dim i As Long, r As ListRow
    For i = 1 To hints.ListRows.Count
        If CStr(hints.DataBodyRange.Cells(i, 1).Value) = msgPath And CStr(hints.DataBodyRange.Cells(i, 2).Value) = ib Then Exit Sub
    Next i
    Set r = hints.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 8).Value = Array(msgPath, ib, material, qty, grDate, vendor, "HIGH", contextText)
    r.Range.Cells(1, 2).NumberFormat = "@"
    r.Range.Cells(1, 2).Value = ib
End Sub
