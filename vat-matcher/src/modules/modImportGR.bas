Attribute VB_Name = "modImportGR"
Option Explicit

Private mTrackingLoadedThisSession As Boolean

Public Function TrackingLoadedThisSession() As Boolean
    TrackingLoadedThisSession = mTrackingLoadedThisSession
End Function

Public Sub PickAndLoadTrackingFile()
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Chon file Theo doi nhap P"
        .Filters.Clear: .Filters.Add "Excel files", "*.xlsx;*.xlsm"
        If .Show <> -1 Then Exit Sub
        LoadTrackingData .SelectedItems(1)
    End With
End Sub

Public Sub LoadTrackingData(ByVal filePath As String)
    Dim src As Workbook, ws As Worksheet, lo As ListObject, hdr As Object
    Dim lastRow As Long, data As Variant, outData() As Variant, i As Long, n As Long
    Dim cReceipt As Long, cMat As Long, cVendor As Long, cDoc As Long, cActual As Long, cIB As Long, cGRDate As Long
    Dim qDoc As Variant, qActual As Variant, flags As String, stage As String
    On Error GoTo EH
    mTrackingLoadedThisSession = False
    Application.ScreenUpdating = False
    stage = "open source": Set src = Workbooks.Open(filePath, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
    stage = "get DATA sheet"
    Set ws = src.Worksheets("DATA")
    stage = "map headers"
    Set hdr = HeaderMap(ws, 27)
    cReceipt = RequiredHeader(hdr, "SO PHIEU")
    cMat = RequiredHeader(hdr, "MA VAT TU")
    cVendor = RequiredHeader(hdr, "NCC")
    cDoc = OptionalHeader(hdr, "QTY CHUNG TU")
    cActual = OptionalHeader(hdr, "QTY THUC TE")
    cIB = OptionalHeader(hdr, "IB")
    cGRDate = RequiredHeader(hdr, "NGAY GR HE THONG")
    If cDoc = 0 And cActual = 0 Then Err.Raise vbObjectError + 100, , "MISSING_REQUIRED_HEADER: Qty document or Qty actual"
    lastRow = Application.Max(ws.Cells(ws.Rows.Count, cReceipt).End(xlUp).Row, ws.Cells(ws.Rows.Count, cMat).End(xlUp).Row)
    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 27)).Value2
    ReDim outData(1 To UBound(data, 1), 1 To 10)
    For i = 1 To UBound(data, 1)
        flags = "": qDoc = Empty: qActual = Empty
        If cDoc > 0 Then qDoc = QtyValue(data(i, cDoc))
        If cActual > 0 Then qActual = QtyValue(data(i, cActual))
        If UseActualQuantity() Then
            If IsEmpty(qActual) Then
                outData(i, 7) = qDoc
                If Not IsEmpty(qDoc) Then flags = "QTY_FALLBACK_DOCUMENT"
            Else
                outData(i, 7) = qActual
            End If
        Else
            If IsEmpty(qDoc) Then
                outData(i, 7) = qActual
                If Not IsEmpty(qActual) Then flags = "QTY_FALLBACK_ACTUAL"
            Else
                outData(i, 7) = qDoc
            End If
        End If
        If Not IsEmpty(qDoc) And Not IsEmpty(qActual) Then
            If Abs(CDbl(qDoc) - CDbl(qActual)) > 0.000001 Then flags = "QTY_DOC_ACTUAL_MISMATCH"
        End If
        outData(i, 1) = i + 1: outData(i, 2) = CStr(data(i, cReceipt)): outData(i, 3) = NormalizeMaterial(CStr(data(i, cMat)))
        outData(i, 4) = NormalizeVendor(CStr(data(i, cVendor))): outData(i, 5) = qDoc: outData(i, 6) = qActual
        If cIB > 0 Then outData(i, 8) = CStr(data(i, cIB))
        outData(i, 9) = data(i, cGRDate): outData(i, 10) = flags
    Next i
    stage = "get destination table": Set lo = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    If lo.ListRows.Count > 0 Then lo.DataBodyRange.Delete
    stage = "write destination"
    lo.Parent.Cells(lo.Range.Row + 1, lo.Range.Column).Resize(UBound(outData, 1), 10).Value = outData
    stage = "resize destination"
    lo.Resize lo.Parent.Range(lo.Range.Cells(1, 1), lo.Parent.Cells(lo.Range.Row + UBound(outData, 1), lo.Range.Column + 9))
    WriteImportLog "INFO", "GR_IMPORT", "Imported " & UBound(outData, 1) & " rows read-only from " & filePath
    'Keep the Tracking P path in a dedicated cell outside tblHome.
    ThisWorkbook.Worksheets("HOME").Range("D4").Value = filePath
    mTrackingLoadedThisSession = True
CleanUp:
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Sub
EH:
    WriteImportLog "ERROR", "GR_IMPORT", stage & ": " & Err.Description
    Resume CleanUp
End Sub

Private Function HeaderMap(ByVal ws As Worksheet, ByVal maxCol As Long) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To maxCol
        d(NormalizeText(CStr(ws.Cells(1, i).Value))) = i
    Next i
    Set HeaderMap = d
End Function
Private Function RequiredHeader(ByVal d As Object, ByVal name As String) As Long
    RequiredHeader = OptionalHeader(d, name)
    If RequiredHeader = 0 Then Err.Raise vbObjectError + 101, , "MISSING_REQUIRED_HEADER: " & name
End Function
Private Function OptionalHeader(ByVal d As Object, ByVal name As String) As Long
    If d.Exists(name) Then OptionalHeader = d(name)
End Function
Private Sub WriteImportLog(ByVal severity As String, ByVal code As String, ByVal message As String)
    LogEvent severity, "LoadTrackingData", code, message
End Sub

Private Function UseActualQuantity() As Boolean
    Dim lo As ListObject, i As Long, configuredValue As String
    On Error GoTo SafeExit
    Set lo = ThisWorkbook.Worksheets("CONFIG").ListObjects("tblConfig")
    If lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(i, 1).Value)), "QtySource", vbTextCompare) = 0 Then
            configuredValue = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(i, 2).Value)))
            UseActualQuantity = (configuredValue = "QTY_ACTUAL" Or InStr(1, configuredValue, "ACTUAL", vbTextCompare) > 0 Or InStr(1, configuredValue, "THUC TE", vbTextCompare) > 0)
            Exit Function
        End If
    Next i
SafeExit:
End Function
