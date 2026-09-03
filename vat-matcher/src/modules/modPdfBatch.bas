Attribute VB_Name = "modPdfBatch"
Option Explicit

Private mPdfFolderSelectedThisSession As Boolean

Public Function PdfFolderSelectedThisSession() As Boolean
    PdfFolderSelectedThisSession = mPdfFolderSelectedThisSession
End Function

Public Sub ConfigurePdfFolderBatch()
    Dim folderPath As String
    On Error GoTo EH
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder containing VAT PDFs"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    ConfigurePdfFolderBatchPath folderPath
    MsgBox "VAT_PDF_BATCH is ready. Refresh the query from HOME or Data > Refresh All.", vbInformation
    Exit Sub
EH:
    LogError "ConfigurePdfFolderBatch", "POWER_QUERY_PDF_UNAVAILABLE", Err.Description, folderPath
    MsgBox "Power Query PDF setup failed. Review LOG.", vbCritical
End Sub

Public Sub ConfigurePdfFolderBatchPath(ByVal folderPath As String)
    Dim formulaText As String
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo EH
    mPdfFolderSelectedThisSession = False
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Err.Raise vbObjectError + 730, , "PDF folder does not exist."
    formulaText = BuildPdfFolderFormula(folderPath)
    On Error Resume Next
    ThisWorkbook.Queries("VAT_PDF_BATCH").Delete
    On Error GoTo EH
    ThisWorkbook.Queries.Add "VAT_PDF_BATCH", formulaText
    LogEvent "INFO", "ConfigurePdfFolderBatch", "PQ_QUERY_READY", "Power Query VAT_PDF_BATCH has been created. Use Data > Refresh All.", folderPath, "Refresh the query and inspect errors before matching."
    LoadPdfQueryToRawSheet
    ThisWorkbook.Worksheets("HOME").Range("B4").Value = folderPath
    mPdfFolderSelectedThisSession = True
    LogEvent "INFO", "ConfigurePdfFolderBatchPath", "PQ_RAW_LOAD_READY", "Query connection is ready in PQ_PDF_RAW.", folderPath, "Use Refresh PDF Batch."
    Exit Sub
EH:
    failureNumber = Err.Number
    failureDescription = Err.Description
    LogError "ConfigurePdfFolderBatchPath", "POWER_QUERY_PDF_UNAVAILABLE", failureDescription, folderPath
    Err.Raise failureNumber, , failureDescription
End Sub

Public Sub RefreshPdfBatch()
    On Error GoTo EH
    ThisWorkbook.Worksheets("PQ_PDF_RAW").ListObjects("tblPdfRaw").QueryTable.Refresh BackgroundQuery:=False
    LogEvent "INFO", "RefreshPdfBatch", "PQ_REFRESH_OK", "PDF batch refresh completed.", CStr(ThisWorkbook.Worksheets("HOME").Range("B4").Value), "Review PDF_FILES and INVOICES before matching."
    Exit Sub
EH:
    LogError "RefreshPdfBatch", "PQ_REFRESH_FAILED", Err.Description, CStr(ThisWorkbook.Worksheets("HOME").Range("B4").Value)
    MsgBox "PDF refresh failed. Review LOG.", vbCritical
End Sub

Private Sub LoadPdfQueryToRawSheet()
    Const CONNECTION_NAME As String = "Query - VAT_PDF_BATCH"
    Const SQL_TEXT As String = "SELECT * FROM [VAT_PDF_BATCH]"
    Dim ws As Worksheet, lo As ListObject, cn As WorkbookConnection
    Dim connectionText As String
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets("PQ_PDF_RAW")
    On Error Resume Next
    ws.ListObjects("tblPdfRaw").Delete
    ThisWorkbook.Connections(CONNECTION_NAME).Delete
    On Error GoTo EH
    connectionText = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=VAT_PDF_BATCH;Extended Properties="""""
    Set cn = ThisWorkbook.Connections.Add(CONNECTION_NAME, "VAT PDF folder batch query", connectionText, SQL_TEXT, xlCmdSql)
    Set lo = ws.ListObjects.Add(xlSrcQuery, cn, True, xlYes, ws.Range("A2"))
    lo.Name = "tblPdfRaw"
    lo.TableStyle = "TableStyleMedium2"
    lo.QueryTable.CommandType = xlCmdSql
    lo.QueryTable.CommandText = SQL_TEXT
    lo.QueryTable.RefreshOnFileOpen = False
    Exit Sub
EH:
    failureNumber = Err.Number
    failureDescription = Err.Description
    LogError "LoadPdfQueryToRawSheet", "PQ_LOAD_BIND_FAILED", failureDescription
    Err.Raise failureNumber, , failureDescription
End Sub

Public Function BuildPdfFolderFormula(ByVal folderPath As String) As String
    Dim q As String
    q = folderPath
    BuildPdfFolderFormula = "let Source = Folder.Files(""" & q & """), PDFs = Table.SelectRows(Source, each Text.Lower([Extension]) = "".pdf""), Parsed = Table.AddColumn(PDFs, ""PdfTables"", each try Pdf.Tables([Content], [Implementation=""1.3""]) otherwise null), Expanded = Table.ExpandTableColumn(Parsed, ""PdfTables"", {""Id"", ""Kind"", ""Name"", ""Data""}, {""PdfTableId"", ""PdfTableKind"", ""PdfTableName"", ""PdfTableData""}), Textified = Table.AddColumn(Expanded, ""PdfText"", each try Text.Combine(List.Transform(Table.ToRows([PdfTableData]), (row) => Text.Combine(List.Transform(row, each try Text.From(_) otherwise """"), "" | "")), "" || "") otherwise """", type text) in Textified"
End Function

Public Sub ResetPdfQueryForRelease()
    Const CONNECTION_NAME As String = "Query - VAT_PDF_BATCH"
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PQ_PDF_RAW")
    ws.ListObjects("tblPdfRaw").Delete
    ThisWorkbook.Connections(CONNECTION_NAME).Delete
    ThisWorkbook.Queries("VAT_PDF_BATCH").Delete
    On Error GoTo 0
    ws.Range("A2:I2").Value = Array("Name", "Extension", "Folder Path", "Date modified", "PdfTableId", "PdfTableKind", "PdfTableName", "PdfTableData", "PdfText")
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A2:I2"), , xlYes)
    lo.Name = "tblPdfRaw"
    lo.TableStyle = "TableStyleMedium2"
End Sub
