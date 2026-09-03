Attribute VB_Name = "modLog"
Option Explicit

Private mCurrentRunId As String

Public Sub SetCurrentRunId(ByVal runId As String)
    mCurrentRunId = Trim$(runId)
End Sub

Public Function CurrentRunId() As String
    If Len(mCurrentRunId) = 0 Then mCurrentRunId = "RUN-" & Format$(Now, "yyyymmdd-hhnnss")
    CurrentRunId = mCurrentRunId
End Function

Public Sub LogEvent(ByVal severity As String, ByVal procedureName As String, ByVal code As String, ByVal message As String, Optional ByVal sourcePath As String = "", Optional ByVal recoveryAction As String = "")
    Dim lo As ListObject, r As ListRow
    On Error GoTo SafeExit
    Set lo = ThisWorkbook.Worksheets("LOG").ListObjects("tblLog")
    Set r = lo.ListRows.Add
    r.Range.Cells(1, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    r.Range.Cells(1, 2).Value = CurrentRunId()
    r.Range.Cells(1, 3).Value = severity
    r.Range.Cells(1, 4).Value = "VAT_Matcher"
    r.Range.Cells(1, 5).Value = procedureName
    r.Range.Cells(1, 6).Value = code
    r.Range.Cells(1, 7).Value = message
    r.Range.Cells(1, 8).Value = sourcePath
    r.Range.Cells(1, 9).Value = recoveryAction
SafeExit:
End Sub

Public Sub LogError(ByVal procedureName As String, ByVal code As String, ByVal message As String, Optional ByVal sourcePath As String = "")
    LogEvent "ERROR", procedureName, code, message, sourcePath, "Review LOG and correct the named stage."
End Sub
