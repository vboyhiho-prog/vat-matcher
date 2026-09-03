Attribute VB_Name = "modPythonBridge"
Option Explicit

' Excel remains the configuration/review UI. This module only exchanges
' normalised UTF-8 CSV files with the portable engine beside the release.
Private Const PY_RUNTIME_FOLDER As String = "vat_python_runtime"
Private Const PORTABLE_ENGINE_RELATIVE_PATH As String = "engine\VAT_Matcher_Engine.exe"

Public Sub RunPythonPipeline()
    Dim runId As String, runRoot As String, inputFolder As String, outputFolder As String
    Dim launcher As String, exitCode As Long
    Dim stdoutPath As String, stderrPath As String
    On Error GoTo EH
    If Not TrackingLoadedThisSession() Then Err.Raise vbObjectError + 980, , "Hay chon va nap lai file P trong phien Excel hien tai."
    If Not PdfFolderSelectedThisSession() Then Err.Raise vbObjectError + 981, , "Hay chon lai thu muc PDF trong phien Excel hien tai."
    If ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData").ListRows.Count = 0 Then Err.Raise vbObjectError + 982, , "File P khong co du lieu de doi soat."
    launcher = ThisWorkbook.Path & "\" & PORTABLE_ENGINE_RELATIVE_PATH
    If Len(Dir$(launcher)) = 0 Then Err.Raise vbObjectError + 983, , "Khong tim thay " & PORTABLE_ENGINE_RELATIVE_PATH & " canh file Excel. Hay dung dung bo release day du; khong can cai Python."
    runId = "PY-" & Format$(Now, "yyyymmdd-hhnnss")
    runRoot = ThisWorkbook.Path & "\" & PY_RUNTIME_FOLDER & "\" & runId
    inputFolder = runRoot & "\input"
    outputFolder = runRoot & "\output"
    stdoutPath = outputFolder & "\engine_stdout.log"
    stderrPath = outputFolder & "\engine_stderr.log"
    EnsureFolderPath inputFolder
    EnsureFolderPath outputFolder
    ExportPythonInput inputFolder, runId
    LogEvent "INFO", "RunPythonPipeline", "PORTABLE_ENGINE_START", "Da xuat file P va cau hinh; dang chay engine PDF dong goi.", PdfFolderPath(), "Engine portable khong can cai Python; sau khi xong xem BC_HOA_DON."
    Application.StatusBar = "VAT Matcher: Engine dang doc PDF va doi soat..."
    exitCode = RunPortableEngine(launcher, inputFolder, outputFolder, stdoutPath, stderrPath)
    Application.StatusBar = False
    If exitCode <> 0 Then Err.Raise vbObjectError + 984, , PortableEngineFailureMessage(exitCode, stderrPath, runRoot)
    If Len(Dir$(outputFolder & "\engine_result.json")) = 0 Then Err.Raise vbObjectError + 985, , "Engine portable khong tao engine_result.json. Kiem tra " & runRoot
    ImportPythonOutput outputFolder
    CreateRenamePreviews
    LogEvent "INFO", "RunPythonPipeline", "PORTABLE_ENGINE_OK", "Da nap ket qua engine portable vao bao cao Excel.", outputFolder, "Review BC_HOA_DON; chi nhap OK sau khi kiem tra."
    Exit Sub
EH:
    Application.StatusBar = False
    LogError "RunPythonPipeline", "PYTHON_ENGINE_FAILED", Err.Description, outputFolder
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function RunPortableEngine(ByVal launcher As String, ByVal inputFolder As String, ByVal outputFolder As String, ByVal stdoutPath As String, ByVal stderrPath As String) As Long
    Dim shell As Object, process As Object, commandLine As String
    Set shell = CreateObject("WScript.Shell")
    commandLine = CmdQuote(launcher) & " --input " & CmdQuote(inputFolder) & " --output " & CmdQuote(outputFolder)
    Set process = shell.Exec(commandLine)
    Do While process.Status = 0
        DoEvents
    Loop
    WriteUtf8Text stdoutPath, process.StdOut.ReadAll
    WriteUtf8Text stderrPath, process.StdErr.ReadAll
    RunPortableEngine = process.ExitCode
End Function

Private Function PortableEngineFailureMessage(ByVal exitCode As Long, ByVal stderrPath As String, ByVal runRoot As String) As String
    Dim detail As String
    If Len(Dir$(stderrPath)) > 0 Then detail = Trim$(ReadUtf8Text(stderrPath))
    If Len(detail) > 600 Then detail = Left$(detail, 600) & "..."
    PortableEngineFailureMessage = "Engine portable ket thuc voi ma " & CStr(exitCode) & ". Kiem tra " & runRoot
    If Len(detail) > 0 Then PortableEngineFailureMessage = PortableEngineFailureMessage & ". Chi tiet: " & detail
End Function

Private Sub ExportPythonInput(ByVal inputFolder As String, ByVal runId As String)
    ExportTableUtf8 ThisWorkbook.Worksheets("CONFIG").ListObjects("tblConfig"), inputFolder & "\config.csv"
    ExportTableUtf8 ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData"), inputFolder & "\gr_data.csv"
    ExportTableUtf8 ThisWorkbook.Worksheets("NCC_MAP").ListObjects("tblVendorMap"), inputFolder & "\ncc_map.csv"
    ExportTableUtf8 ThisWorkbook.Worksheets("PARSER_PROFILES").ListObjects("tblParserProfiles"), inputFolder & "\parser_profiles.csv"
    ExportTableUtf8 ThisWorkbook.Worksheets("MATERIAL_SCOPE_MAP").ListObjects("tblMaterialScopeMap"), inputFolder & "\material_scope_map.csv"
    ExportTableUtf8 ThisWorkbook.Worksheets("MATERIAL_NCC_MAP").ListObjects("tblMaterialVendorMap"), inputFolder & "\material_ncc_map.csv"
    WriteUtf8Text inputFolder & "\run_config.json", "{""run_id"":""" & JsonEscape(runId) & """,""pdf_folder"":""" & JsonEscape(PdfFolderPath()) & """,""workbook"":""" & JsonEscape(ThisWorkbook.Name) & """}"
End Sub

Private Sub ImportPythonOutput(ByVal outputFolder As String)
    ImportCsvIntoTable outputFolder & "\pdf_files.csv", ThisWorkbook.Worksheets("PDF_FILES").ListObjects("tblPdfFiles")
    ImportCsvIntoTable outputFolder & "\invoices.csv", ThisWorkbook.Worksheets("INVOICES").ListObjects("tblInvoices")
    ImportCsvIntoTable outputFolder & "\vat_lines.csv", ThisWorkbook.Worksheets("VAT_LINES").ListObjects("tblVatLines")
    ImportCsvIntoTable outputFolder & "\match_candidates.csv", ThisWorkbook.Worksheets("MATCH_CANDIDATES").ListObjects("tblCandidates")
    ImportCsvIntoTable outputFolder & "\allocations.csv", ThisWorkbook.Worksheets("ALLOCATIONS").ListObjects("tblAllocations")
    ImportCsvIntoTable outputFolder & "\invoice_report.csv", ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    ImportCsvIntoTable outputFolder & "\receipt_report.csv", ThisWorkbook.Worksheets("BC_PHIEU").ListObjects("tblReceiptReport")
    ImportCsvIntoTable outputFolder & "\engine_log.csv", ThisWorkbook.Worksheets("LOG").ListObjects("tblLog")
    ConfigureWorkbookUsability
End Sub

Private Sub ExportTableUtf8(ByVal lo As ListObject, ByVal filePath As String)
    Dim content As String, headers As Variant, data As Variant, i As Long, j As Long, lineText As String
    headers = lo.HeaderRowRange.Value2
    For j = 1 To UBound(headers, 2)
        If j > 1 Then lineText = lineText & ","
        lineText = lineText & CsvEscape(CStr(headers(1, j)))
    Next j
    content = lineText & vbCrLf
    If Not lo.DataBodyRange Is Nothing Then
        data = lo.DataBodyRange.Value2
        For i = 1 To UBound(data, 1)
            lineText = ""
            For j = 1 To UBound(data, 2)
                If j > 1 Then lineText = lineText & ","
                lineText = lineText & CsvEscape(ExportValue(data(i, j)))
            Next j
            content = content & lineText & vbCrLf
        Next i
    End If
    WriteUtf8Text filePath, content
End Sub

Private Function ExportValue(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then Exit Function
    'Value2 returns Excel dates as serial numbers. Keep that raw typed value;
    'Python recognises the serial and this avoids mistaking a quantity such as
    '1000 for a date under a locale-specific IsDate conversion.
    ExportValue = CStr(value)
End Function

Private Function CsvEscape(ByVal value As String) As String
    value = Replace(Replace(value, vbCr, " "), vbLf, " ")
    CsvEscape = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Sub ImportCsvIntoTable(ByVal filePath As String, ByVal lo As ListObject)
    Dim textData As String, rawLines As Variant, header As Variant, fields As Variant
    Dim map As Object, values() As Variant, i As Long, j As Long, rowCount As Long, fieldIndex As Long
    If Len(Dir$(filePath)) = 0 Then Err.Raise vbObjectError + 986, , "Thieu file ket qua Python: " & filePath
    textData = Replace(ReadUtf8Text(filePath), vbCrLf, vbLf)
    textData = Replace(textData, vbCr, vbLf)
    rawLines = Split(textData, vbLf)
    If UBound(rawLines) < 0 Or Len(CStr(rawLines(0))) = 0 Then Err.Raise vbObjectError + 987, , "CSV Python rong: " & filePath
    header = ParseCsvLine(CStr(rawLines(0)))
    Set map = CreateObject("Scripting.Dictionary")
    For j = LBound(header) To UBound(header): map(CStr(header(j))) = j: Next j
    For i = 1 To UBound(rawLines)
        If Len(CStr(rawLines(i))) > 0 Then rowCount = rowCount + 1
    Next i
    ClearOutputTable lo
    If rowCount = 0 Then Exit Sub
    ReDim values(1 To rowCount, 1 To lo.ListColumns.Count)
    rowCount = 0
    For i = 1 To UBound(rawLines)
        If Len(CStr(rawLines(i))) > 0 Then
            rowCount = rowCount + 1
            fields = ParseCsvLine(CStr(rawLines(i)))
            For j = 1 To lo.ListColumns.Count
                If map.Exists(lo.ListColumns(j).Name) Then
                    fieldIndex = CLng(map(lo.ListColumns(j).Name))
                    If fieldIndex <= UBound(fields) Then values(rowCount, j) = ImportValue(CStr(fields(fieldIndex)), lo.ListColumns(j).Name)
                End If
            Next j
        End If
    Next i
    lo.Parent.Cells(lo.Range.Row + 1, lo.Range.Column).Resize(UBound(values, 1), UBound(values, 2)).Value = values
    lo.Resize lo.Parent.Range(lo.Range.Cells(1, 1), lo.Parent.Cells(lo.Range.Row + UBound(values, 1), lo.Range.Column + lo.ListColumns.Count - 1))
End Sub

Private Function ImportValue(ByVal value As String, ByVal columnName As String) As Variant
    If columnName = "InvoiceDate" Or columnName = "GRDate" Or columnName = "Timestamp" Then
        If Len(value) > 0 Then
            If IsDate(value) Then ImportValue = CDate(value): Exit Function
        End If
    End If
    ImportValue = value
End Function

Private Function ParseCsvLine(ByVal lineText As String) As Variant
    Dim values() As String, currentValue As String, i As Long, inQuotes As Boolean, charValue As String, nextValue As String, count As Long
    ReDim values(0 To 0)
    For i = 1 To Len(lineText)
        charValue = Mid$(lineText, i, 1)
        If charValue = """" Then
            If inQuotes And i < Len(lineText) Then
                nextValue = Mid$(lineText, i + 1, 1)
                If nextValue = """" Then currentValue = currentValue & """": i = i + 1 Else inQuotes = False
            Else
                inQuotes = True
            End If
        ElseIf charValue = "," And Not inQuotes Then
            values(count) = currentValue: count = count + 1: ReDim Preserve values(0 To count): currentValue = ""
        Else
            currentValue = currentValue & charValue
        End If
    Next i
    values(count) = currentValue
    ParseCsvLine = values
End Function

Private Sub ClearOutputTable(ByVal lo As ListObject)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
End Sub

Private Sub EnsureFolderPath(ByVal folderPath As String)
    Dim fso As Object, parentPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then Exit Sub
    parentPath = fso.GetParentFolderName(folderPath)
    If Not fso.FolderExists(parentPath) Then EnsureFolderPath parentPath
    fso.CreateFolder folderPath
End Sub

Private Sub WriteUtf8Text(ByVal filePath As String, ByVal content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

Private Function ReadUtf8Text(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile filePath
    ReadUtf8Text = stream.ReadText
    stream.Close
End Function

Private Function CmdQuote(ByVal value As String) As String
    CmdQuote = Chr$(34) & value & Chr$(34)
End Function

Private Function JsonEscape(ByVal value As String) As String
    value = Replace(value, Chr$(92), Chr$(92) & Chr$(92))
    value = Replace(value, Chr$(34), Chr$(92) & Chr$(34))
    value = Replace(value, vbCr, Chr$(92) & "r")
    JsonEscape = Replace(value, vbLf, Chr$(92) & "n")
End Function
