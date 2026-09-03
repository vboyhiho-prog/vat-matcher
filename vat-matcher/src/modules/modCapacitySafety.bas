Attribute VB_Name = "modCapacitySafety"
Option Explicit

Public Sub ApplyAllocationSafetyPass()
    Dim lines As ListObject, allocations As ListObject, report As ListObject
    Dim i As Long, invoiceId As String, demand As Double, allocated As Double, reportRow As Long, existingReason As String, material As String, issueCount As Long
    On Error GoTo EH
    Set lines = ThisWorkbook.Worksheets("VAT_LINES").ListObjects("tblVatLines")
    Set allocations = ThisWorkbook.Worksheets("ALLOCATIONS").ListObjects("tblAllocations")
    Set report = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    MarkGlobalCapacityConflicts allocations, report
    For i = 1 To lines.ListRows.Count
        invoiceId = CStr(lines.DataBodyRange.Cells(i, 2).Value)
        material = CStr(lines.DataBodyRange.Cells(i, 5).Value)
        demand = CDbl(Val(CStr(lines.DataBodyRange.Cells(i, 7).Value)))
        allocated = AllocationForVatLine(allocations, CStr(lines.DataBodyRange.Cells(i, 1).Value))
        reportRow = ReportRowForInvoice(report, Replace(invoiceId, "INV-", ""))
        If reportRow > 0 Then
            If CapacityStatus(CStr(report.DataBodyRange.Cells(reportRow, 8).Value), allocated, demand) = "SUSPECT_CONFLICT" Then
                report.DataBodyRange.Cells(reportRow, 8).Value = "SUSPECT_CONFLICT"
                existingReason = CStr(report.DataBodyRange.Cells(reportRow, 11).Value)
                If InStr(1, existingReason, "Tat ca ma vat tu tren hoa don da khop", vbTextCompare) = 1 Then existingReason = ""
                issueCount = CountCapacityIssues(existingReason)
                If issueCount >= 4 Then
                    If InStr(1, existingReason, "Con ma vat tu khac khong du so luong", vbTextCompare) = 0 Then existingReason = existingReason & " | Con ma vat tu khac khong du so luong; xem VAT_LINES."
                ElseIf InStr(1, existingReason, "Ma " & material & " khong du so luong", vbTextCompare) = 0 Then
                    If Len(existingReason) > 0 Then existingReason = existingReason & " | "
                    existingReason = existingReason & "Ma " & material & " khong du so luong tren phieu de xuat."
                End If
                report.DataBodyRange.Cells(reportRow, 11).Value = existingReason
            End If
        End If
    Next i
    LogEvent "INFO", "ApplyAllocationSafetyPass", "CAPACITY_SAFETY_DONE", "Allocation residuals checked.", "", "Review SUSPECT_CONFLICT before any decision."
    Exit Sub
EH:
    LogError "ApplyAllocationSafetyPass", "CAPACITY_SAFETY_FAILED", Err.Description
End Sub

Public Function GlobalCapacityConflictCount() As Long
    Dim gr As ListObject, allocations As ListObject, capacities As Object, totals As Object
    Dim grData As Variant, allocationData As Variant, i As Long, sourceKey As String, sourceItem As Variant
    Set gr = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    Set allocations = ThisWorkbook.Worksheets("ALLOCATIONS").ListObjects("tblAllocations")
    Set capacities = CreateObject("Scripting.Dictionary")
    Set totals = CreateObject("Scripting.Dictionary")
    If Not gr.DataBodyRange Is Nothing Then
        grData = gr.DataBodyRange.Value2
        For i = 1 To UBound(grData, 1)
            sourceKey = CStr(grData(i, 1))
            capacities(sourceKey) = CDbl(Val(CStr(grData(i, 7))))
        Next i
    End If
    If Not allocations.DataBodyRange Is Nothing Then
        allocationData = allocations.DataBodyRange.Value2
        For i = 1 To UBound(allocationData, 1)
            sourceKey = CStr(allocationData(i, 5))
            If totals.Exists(sourceKey) Then totals(sourceKey) = CDbl(totals(sourceKey)) + CDbl(Val(CStr(allocationData(i, 6)))) Else totals.Add sourceKey, CDbl(Val(CStr(allocationData(i, 6))))
        Next i
    End If
    For Each sourceItem In totals.Keys
        If Not capacities.Exists(CStr(sourceItem)) Or CDbl(totals(sourceItem)) > CDbl(capacities(sourceItem)) + 0.000001 Then GlobalCapacityConflictCount = GlobalCapacityConflictCount + 1
    Next sourceItem
End Function

Public Function GlobalSourceRowMultiInvoiceCount() As Long
    Dim allocations As ListObject, owners As Object, conflicts As Object, allocationData As Variant
    Dim i As Long, sourceKey As String, invoiceId As String, sourceItem As Variant
    Set allocations = ThisWorkbook.Worksheets("ALLOCATIONS").ListObjects("tblAllocations")
    Set owners = CreateObject("Scripting.Dictionary")
    Set conflicts = CreateObject("Scripting.Dictionary")
    If allocations.DataBodyRange Is Nothing Then Exit Function
    allocationData = allocations.DataBodyRange.Value2
    For i = 1 To UBound(allocationData, 1)
        sourceKey = CStr(allocationData(i, 5))
        invoiceId = CStr(allocationData(i, 2))
        If owners.Exists(sourceKey) Then
            If CStr(owners(sourceKey)) <> invoiceId Then conflicts(sourceKey) = True
        Else
            owners.Add sourceKey, invoiceId
        End If
    Next i
    For Each sourceItem In conflicts.Keys
        GlobalSourceRowMultiInvoiceCount = GlobalSourceRowMultiInvoiceCount + 1
    Next sourceItem
End Function

Public Function AllocationUsesMismatchFlagCount() As Long
    Dim gr As ListObject, allocations As ListObject, flagged As Object, grData As Variant, allocationData As Variant, i As Long, sourceKey As String
    Set gr = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    Set allocations = ThisWorkbook.Worksheets("ALLOCATIONS").ListObjects("tblAllocations")
    Set flagged = CreateObject("Scripting.Dictionary")
    If Not gr.DataBodyRange Is Nothing Then
        grData = gr.DataBodyRange.Value2
        For i = 1 To UBound(grData, 1)
            If InStr(1, UCase$(CStr(grData(i, 10))), "QTY_DOC_ACTUAL_MISMATCH", vbTextCompare) > 0 Then flagged(CStr(grData(i, 1))) = True
        Next i
    End If
    If Not allocations.DataBodyRange Is Nothing Then
        allocationData = allocations.DataBodyRange.Value2
        For i = 1 To UBound(allocationData, 1)
            sourceKey = CStr(allocationData(i, 5))
            If flagged.Exists(sourceKey) Then AllocationUsesMismatchFlagCount = AllocationUsesMismatchFlagCount + 1
        Next i
    End If
End Function

Private Sub MarkGlobalCapacityConflicts(ByVal allocations As ListObject, ByVal report As ListObject)
    Dim gr As ListObject, capacities As Object, totals As Object, conflicts As Object
    Dim grData As Variant, allocationData As Variant, i As Long, sourceKey As String, invoiceId As String, sourceItem As Variant
    Set gr = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    Set capacities = CreateObject("Scripting.Dictionary")
    Set totals = CreateObject("Scripting.Dictionary")
    Set conflicts = CreateObject("Scripting.Dictionary")
    If Not gr.DataBodyRange Is Nothing Then
        grData = gr.DataBodyRange.Value2
        For i = 1 To UBound(grData, 1)
            capacities(CStr(grData(i, 1))) = CDbl(Val(CStr(grData(i, 7))))
        Next i
    End If
    If Not allocations.DataBodyRange Is Nothing Then
        allocationData = allocations.DataBodyRange.Value2
        For i = 1 To UBound(allocationData, 1)
            sourceKey = CStr(allocationData(i, 5))
            If totals.Exists(sourceKey) Then totals(sourceKey) = CDbl(totals(sourceKey)) + CDbl(Val(CStr(allocationData(i, 6)))) Else totals.Add sourceKey, CDbl(Val(CStr(allocationData(i, 6))))
        Next i
    End If
    For Each sourceItem In totals.Keys
        If Not capacities.Exists(CStr(sourceItem)) Or CDbl(totals(sourceItem)) > CDbl(capacities(sourceItem)) + 0.000001 Then conflicts(CStr(sourceItem)) = True
    Next sourceItem
    If conflicts.Count = 0 Or allocations.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To UBound(allocationData, 1)
        sourceKey = CStr(allocationData(i, 5))
        If conflicts.Exists(sourceKey) Then
            allocations.DataBodyRange.Cells(i, 8).Value = "CAPACITY_CONFLICT"
            invoiceId = CStr(allocationData(i, 2))
            MarkReportCapacityConflict report, invoiceId, sourceKey
        End If
    Next i
End Sub

Private Sub MarkReportCapacityConflict(ByVal report As ListObject, ByVal invoiceId As String, ByVal sourceKey As String)
    Dim reportRow As Long, currentNote As String
    reportRow = ReportRowForInvoice(report, Replace(invoiceId, "INV-", ""))
    If reportRow = 0 Then Exit Sub
    report.DataBodyRange.Cells(reportRow, 8).Value = "SUSPECT_CONFLICT"
    currentNote = CStr(report.DataBodyRange.Cells(reportRow, 11).Value)
    If InStr(1, currentNote, "CAPACITY_CONFLICT", vbTextCompare) = 0 Then
        If Len(currentNote) > 0 Then currentNote = currentNote & " | "
        report.DataBodyRange.Cells(reportRow, 11).Value = currentNote & "CAPACITY_CONFLICT: SourceRow " & sourceKey & " bi phan bo vuot so luong."
    End If
End Sub

Private Function CountCapacityIssues(ByVal noteText As String) As Long
    Dim pos As Long
    pos = InStr(1, noteText, "khong du so luong tren phieu", vbTextCompare)
    Do While pos > 0
        CountCapacityIssues = CountCapacityIssues + 1
        pos = InStr(pos + 1, noteText, "khong du so luong tren phieu", vbTextCompare)
    Loop
End Function

Public Function CapacityStatus(ByVal currentStatus As String, ByVal allocated As Double, ByVal demand As Double) As String
    CapacityStatus = currentStatus
    If currentStatus = "MATCHED" And allocated + 0.000001 < demand Then CapacityStatus = "SUSPECT_CONFLICT"
End Function

Private Function AllocationForVatLine(ByVal allocations As ListObject, ByVal vatLineId As String) As Double
    Dim i As Long
    For i = 1 To allocations.ListRows.Count
        If CStr(allocations.DataBodyRange.Cells(i, 3).Value) = vatLineId Then AllocationForVatLine = AllocationForVatLine + CDbl(Val(CStr(allocations.DataBodyRange.Cells(i, 6).Value)))
    Next i
End Function

Private Function ReportRowForInvoice(ByVal report As ListObject, ByVal invoiceNo As String) As Long
    Dim i As Long
    For i = 1 To report.ListRows.Count
        If CStr(report.DataBodyRange.Cells(i, 3).Value) = invoiceNo Then ReportRowForInvoice = i: Exit Function
    Next i
End Function
