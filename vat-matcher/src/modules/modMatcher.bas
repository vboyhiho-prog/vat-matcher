Attribute VB_Name = "modMatcher"
Option Explicit

Private mAutomationMode As Boolean
Private mReceiptEmailHint As Object
Private mInScopeMaterials As Object
Private mOutOfScopeMaterials As Object
Private mMaterialVendors As Object
Private mAllReceipts As Object
Private Const MAX_RECEIPT_CANDIDATES As Long = 5000

Public Sub SetMatcherAutomationMode(ByVal enabled As Boolean)
    mAutomationMode = enabled
End Sub

Public Sub VietHoaLyDoTrongBaoCao()
    Dim report As ListObject, i As Long
    Set report = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    If report.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To report.ListRows.Count
        report.DataBodyRange.Cells(i, 11).Value = VietHoaLyDo(CStr(report.DataBodyRange.Cells(i, 11).Value))
    Next i
End Sub

Public Sub RunMatch()
    Dim invoices As ListObject, vatLines As ListObject, grData As ListObject
    Dim candidates As ListObject, allocations As ListObject, invoiceReport As ListObject
    Dim invData As Variant, lineData As Variant, gr As Variant
    Dim receiptMaterialQty As Object, vendorReceipts As Object, receiptDates As Object, materialReceipts As Object, grRowsByBucket As Object
    Dim usedQty As Object, usedReceiptMaterialQty As Object, materialVendors As Object
    Dim i As Long, j As Long, lineCount As Long, matchedLines As Long, searchCount As Long, inScopeCount As Long, outScopeCount As Long, unknownCount As Long
    Dim invoiceId As String, vendor As String, receipt As Variant, material As String
    Dim totalQtyScore As Double, coverage As Double, qtyScore As Double, dateScore As Double, totalScore As Double, bestReceipt As String, bestScore As Double, bestPoints As Double
    Dim candidateId As String, runId As String, status As String, hintReceiptSet As String, reportReason As String
    Dim exactAssignments As Object
    Dim searchTruncated As Boolean, dateEligible As Double, demandLineCount As Long
    On Error GoTo EH
    If Not mAutomationMode Then
        If Not TrackingLoadedThisSession() Then Err.Raise vbObjectError + 811, , "Chon va nap lai file P trong phien Excel hien tai truoc khi khop."
        If Not PdfFolderSelectedThisSession() Then Err.Raise vbObjectError + 812, , "Chon lai thu muc PDF trong phien Excel hien tai truoc khi khop."
    End If
    Set invoices = ThisWorkbook.Worksheets("INVOICES").ListObjects("tblInvoices")
    Set vatLines = ThisWorkbook.Worksheets("VAT_LINES").ListObjects("tblVatLines")
    Set grData = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    Set candidates = ThisWorkbook.Worksheets("MATCH_CANDIDATES").ListObjects("tblCandidates")
    Set allocations = ThisWorkbook.Worksheets("ALLOCATIONS").ListObjects("tblAllocations")
    Set invoiceReport = ThisWorkbook.Worksheets("BC_HOA_DON").ListObjects("tblInvoiceReport")
    If invoices.DataBodyRange Is Nothing Or vatLines.DataBodyRange Is Nothing Or grData.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 810, , "INVOICES, VAT_LINES, and GR_DATA must contain data."
    invData = invoices.DataBodyRange.Value2
    lineData = vatLines.DataBodyRange.Value2
    gr = grData.DataBodyRange.Value2
    Set materialVendors = CreateObject("Scripting.Dictionary")
    BuildMaterialVendorIndex invoices, vatLines, materialVendors
    CanonicalizeGrVendorsByMaterial gr, materialVendors
    Set mMaterialVendors = materialVendors
    Set receiptMaterialQty = CreateObject("Scripting.Dictionary")
    Set vendorReceipts = CreateObject("Scripting.Dictionary")
    Set receiptDates = CreateObject("Scripting.Dictionary")
    Set materialReceipts = CreateObject("Scripting.Dictionary")
    Set grRowsByBucket = CreateObject("Scripting.Dictionary")
    Set mAllReceipts = CreateObject("Scripting.Dictionary")
    BuildReceiptIndexes gr, receiptMaterialQty, vendorReceipts, receiptDates, materialReceipts, grRowsByBucket
    Set mInScopeMaterials = CreateObject("Scripting.Dictionary")
    BuildInScopeMaterialIndex gr, mInScopeMaterials
    Set mOutOfScopeMaterials = CreateObject("Scripting.Dictionary")
    BuildOutOfScopeMaterialIndex mOutOfScopeMaterials
    ClassifyVatLines vatLines
    ClearRows candidates
    ClearRows allocations
    ClearRows invoiceReport
    Set usedQty = CreateObject("Scripting.Dictionary")
    Set usedReceiptMaterialQty = CreateObject("Scripting.Dictionary")
    Set mReceiptEmailHint = Nothing
    runId = "RUN-" & Format$(Now, "yyyymmdd-hhnnss")
    SetCurrentRunId runId
    For i = 1 To UBound(invData, 1)
        invoiceId = CStr(invData(i, 1))
        vendor = CStr(invData(i, 7))
        bestReceipt = ""
        bestScore = -1
        bestPoints = 0
        searchCount = 0
        searchTruncated = False
        InvoiceScopeCounts lineData, invoiceId, inScopeCount, outScopeCount, unknownCount
        If inScopeCount = 0 And unknownCount = 0 Then
            AddInvoiceReport invoiceReport, runId, invData, i, "", 0, "OTHER_FACTORY", "NO_IN_SCOPE_MATERIAL: all " & CStr(outScopeCount) & " VAT line(s) are outside the GR snapshot."
            LogEvent "INFO", "RunMatch", "OTHER_FACTORY", CStr(invData(i, 5)) & " has no in-scope material.", "", "Eligible only for controlled VAT XUONG KHAC preview."
            GoTo NextInvoice
        End If
        If vendorReceipts.Exists(vendor) Then
            For Each receipt In vendorReceipts(vendor).Keys
                dateEligible = DateEligibilityForReceipt(invData(i, 6), vendor, CStr(receipt), receiptDates)
                If dateEligible > 0 Then
                    searchCount = searchCount + 1
                    If searchCount > MAX_RECEIPT_CANDIDATES Then searchTruncated = True: Exit For
                    ScoreReceipt lineData, invoiceId, vendor, CStr(receipt), receiptMaterialQty, lineCount, matchedLines, totalQtyScore
                Else
                    lineCount = 0: matchedLines = 0: totalQtyScore = 0
                End If
                If lineCount > 0 And matchedLines > 0 Then
                    coverage = matchedLines / lineCount
                    qtyScore = totalQtyScore / lineCount
                    dateScore = DatePointsForReceipt(invData(i, 6), vendor, CStr(receipt), receiptDates)
                    totalScore = 20 + (coverage * 40) + (qtyScore * 20) + dateScore + EmailHintPoints(CStr(receipt))
                    candidateId = "C-" & CStr(i) & "-" & CStr(receipt)
                    AddCandidate candidates, candidateId, invoiceId, CStr(receipt), 20, coverage * 40, qtyScore * 20, dateScore, EmailHintPoints(CStr(receipt)), totalScore, 1, CandidateReason(matchedLines, lineCount, dateScore, EmailHintPoints(CStr(receipt)))
                    If (coverage * 0.8) + (DateEligibilityForReceipt(invData(i, 6), vendor, CStr(receipt), receiptDates) * 0.2) > bestScore Then
                        bestScore = (coverage * 0.8) + (DateEligibilityForReceipt(invData(i, 6), vendor, CStr(receipt), receiptDates) * 0.2)
                        bestReceipt = CStr(receipt)
                        bestPoints = totalScore
                    End If
                End If
            Next receipt
        End If
        hintReceiptSet = ActiveReceiptHint(CStr(invData(i, 5)))
        If Not searchTruncated And Len(hintReceiptSet) > 0 Then
            AddCandidate candidates, "C-" & CStr(i) & "-HINT", invoiceId, hintReceiptSet, 0, 0, 0, 0, 0, 0, 1, "Reference hint only; quantity and date must still match."
        End If
        Set exactAssignments = BuildExactAssignments(lineData, gr, invoiceId, vendor, invData(i, 6), receiptDates, materialReceipts, grRowsByBucket, usedQty, hintReceiptSet)
        demandLineCount = InvoiceDemandLineCount(lineData, invoiceId)
        If exactAssignments.Count > 0 Then
            bestReceipt = ReceiptSetFromAssignments(exactAssignments)
            bestScore = exactAssignments.Count / demandLineCount
            bestPoints = Round(bestScore * 100, 1)
            AddCandidate candidates, "C-" & CStr(i) & "-ALLOC", invoiceId, bestReceipt, 20, bestScore * 40, bestScore * 20, 0, EmailHintPoints(bestReceipt), bestPoints, 1, "Allocation-safe exact coverage " & CStr(exactAssignments.Count) & "/" & CStr(demandLineCount)
        End If
        reportReason = ""
        If searchTruncated Then
            bestReceipt = ""
            bestPoints = 0
            status = "SUSPECT"
            reportReason = "SEARCH_TRUNCATED: more than " & CStr(MAX_RECEIPT_CANDIDATES) & " receipt candidates; refine vendor/date filters."
            AddCandidate candidates, "C-" & CStr(i) & "-LIMIT", invoiceId, "", 0, 0, 0, 0, 0, 0, 0, reportReason
        ElseIf exactAssignments.Count > 0 Then
            If exactAssignments.Count = demandLineCount Then
                status = "MATCHED"
            Else
                status = "PARTIAL_MATCHED"
                reportReason = "PARTIAL_REVIEW: da khop " & CStr(exactAssignments.Count) & "/" & CStr(demandLineCount) & " dong vat tu; diem " & Format$(bestPoints, "0.0") & "%. Can review; chi doi ten khi nguoi dung tu nhap OK."
            End If
            AllocateExactAssignments lineData, gr, invoiceId, vendor, exactAssignments, grRowsByBucket, usedQty, usedReceiptMaterialQty, allocations
        Else
            status = "SUSPECT"
        End If
        If outScopeCount > 0 Then
            If Len(reportReason) > 0 Then reportReason = reportReason & " | "
            reportReason = reportReason & "MIXED_SCOPE: " & CStr(outScopeCount) & " out-of-scope VAT line(s) ignored for allocation."
        End If
        reportReason = BuildInvoiceNote(lineData, invoiceId, vendor, bestReceipt, materialReceipts, receiptDates, invData(i, 6), status, reportReason)
        AddInvoiceReport invoiceReport, runId, invData, i, bestReceipt, bestPoints, status, reportReason
        RankCandidatesForInvoice candidates, invoiceId
NextInvoice:
    Next i
    ApplyAllocationSafetyPass
    LogEvent "INFO", "RunMatch", "MATCH_OK", CStr(invoiceReport.ListRows.Count) & " invoices scored; " & CStr(allocations.ListRows.Count) & " allocations created.", "", "Review BC_HOA_DON and manual decisions before rename."
    Exit Sub
EH:
    LogError "RunMatch", "MATCH_FAILED", Err.Description
    If Not mAutomationMode Then MsgBox "Matching failed. Review LOG.", vbCritical
End Sub

Private Sub BuildInScopeMaterialIndex(ByVal gr As Variant, ByVal materials As Object)
    Dim i As Long, material As String
    For i = 1 To UBound(gr, 1)
        material = CStr(gr(i, 3))
        If Len(material) > 0 Then materials(material) = True
    Next i
End Sub

Private Sub ClassifyVatLines(ByVal lines As ListObject)
    Dim i As Long, scopeColumn As Long
    On Error Resume Next
    scopeColumn = lines.ListColumns("ScopeStatus").Index
    On Error GoTo 0
    If scopeColumn = 0 Or lines.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lines.ListRows.Count
        lines.DataBodyRange.Cells(i, scopeColumn).Value = MaterialScope(CStr(lines.DataBodyRange.Cells(i, 5).Value))
    Next i
End Sub

Private Sub InvoiceScopeCounts(ByVal lineData As Variant, ByVal invoiceId As String, ByRef inScopeCount As Long, ByRef outScopeCount As Long, ByRef unknownCount As Long)
    Dim i As Long, scope As String
    inScopeCount = 0: outScopeCount = 0: unknownCount = 0
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) = invoiceId Then
            scope = MaterialScope(CStr(lineData(i, 5)))
            If scope = "IN_SCOPE" Then
                inScopeCount = inScopeCount + 1
            ElseIf scope = "OUT_OF_SCOPE_MATERIAL" Then
                outScopeCount = outScopeCount + 1
            Else
                unknownCount = unknownCount + 1
            End If
        End If
    Next i
End Sub

Private Function IsInScopeMaterial(ByVal material As String) As Boolean
    If mInScopeMaterials Is Nothing Then Exit Function
    IsInScopeMaterial = mInScopeMaterials.Exists(material)
End Function

Private Function IsDemandMaterial(ByVal material As String) As Boolean
    IsDemandMaterial = (MaterialScope(material) <> "OUT_OF_SCOPE_MATERIAL")
End Function

Private Function MaterialScope(ByVal material As String) As String
    If IsInScopeMaterial(material) Then
        MaterialScope = "IN_SCOPE"
    ElseIf Not mOutOfScopeMaterials Is Nothing And mOutOfScopeMaterials.Exists(material) Then
        MaterialScope = "OUT_OF_SCOPE_MATERIAL"
    Else
        MaterialScope = "UNKNOWN_MATERIAL"
    End If
End Function

Private Sub BuildOutOfScopeMaterialIndex(ByVal materials As Object)
    Dim lo As ListObject, i As Long
    On Error GoTo SafeExit
    Set lo = ThisWorkbook.Worksheets("MATERIAL_SCOPE_MAP").ListObjects("tblMaterialScopeMap")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If UCase$(CStr(lo.DataBodyRange.Cells(i, 2).Value)) = "OUT_OF_SCOPE_MATERIAL" Then materials(CStr(lo.DataBodyRange.Cells(i, 1).Value)) = True
    Next i
SafeExit:
End Sub

Private Function ActiveReceiptHint(ByVal invoiceNo As String) As String
    Dim lo As ListObject, i As Long, candidate As String, receipt As Variant
    On Error GoTo SafeExit
    Set lo = ThisWorkbook.Worksheets("MATCH_HINTS").ListObjects("tblMatchHints")
    For i = 1 To lo.ListRows.Count
        If Right$("00000000" & CStr(lo.DataBodyRange.Cells(i, 1).Value), 8) = invoiceNo And CBool(lo.DataBodyRange.Cells(i, 4).Value) Then
            candidate = CStr(lo.DataBodyRange.Cells(i, 2).Value)
            For Each receipt In Split(candidate, "+")
                If Not ReceiptExistsInGr(CStr(receipt)) Then Exit Function
            Next receipt
            ActiveReceiptHint = candidate
            Exit Function
        End If
    Next i
SafeExit:
End Function

Private Function ReceiptExistsInGr(ByVal receiptNo As String) As Boolean
    If mAllReceipts Is Nothing Then Exit Function
    ReceiptExistsInGr = mAllReceipts.Exists(receiptNo)
End Function

Private Function BestReceiptSet(ByVal lineData As Variant, ByVal invoiceId As String, ByVal vendor As String, ByVal receipts As Object, ByVal receiptMaterialQty As Object, ByVal invoiceDate As Variant, ByVal receiptDates As Object) As Object
    Dim selected As Object, receipt As Variant, bestReceipt As String, bestCoverage As Double, coverage As Double
    Dim lineCount As Long, matchedLines As Long, totalQtyScore As Double, pass As Long
    Set selected = CreateObject("Scripting.Dictionary")
    For pass = 1 To 3
        bestReceipt = "": bestCoverage = -1
        For Each receipt In receipts.Keys
            If Not selected.Exists(CStr(receipt)) And DateEligibilityForReceipt(invoiceDate, vendor, CStr(receipt), receiptDates) > 0 Then
                selected.Add CStr(receipt), True
                ScoreReceiptSet lineData, invoiceId, vendor, selected, receiptMaterialQty, lineCount, matchedLines, totalQtyScore
                coverage = matchedLines / lineCount
                selected.Remove CStr(receipt)
                If coverage > bestCoverage Then bestCoverage = coverage: bestReceipt = CStr(receipt)
            End If
        Next receipt
        If Len(bestReceipt) = 0 Then Exit For
        selected.Add bestReceipt, True
        If bestCoverage >= 0.999 Then Exit For
    Next pass
    Set BestReceiptSet = selected
End Function

Private Sub ScoreReceiptSet(ByVal lineData As Variant, ByVal invoiceId As String, ByVal vendor As String, ByVal receiptSet As Object, ByVal receiptMaterialQty As Object, ByRef lineCount As Long, ByRef matchedLines As Long, ByRef totalQtyScore As Double)
    Dim i As Long, receipt As Variant, material As String, requiredQty As Double, availableQty As Double, key As String
    lineCount = 0: matchedLines = 0: totalQtyScore = 0
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) = invoiceId And IsDemandMaterial(CStr(lineData(i, 5))) Then
            lineCount = lineCount + 1
            material = CStr(lineData(i, 5))
            requiredQty = CDbl(Val(CStr(lineData(i, 7))))
            availableQty = 0
            For Each receipt In receiptSet.Keys
                key = vendor & "|" & CStr(receipt) & "|" & material
                If receiptMaterialQty.Exists(key) Then availableQty = availableQty + CDbl(receiptMaterialQty(key))
            Next receipt
            If Abs(availableQty - requiredQty) < 0.0001 Then
                matchedLines = matchedLines + 1
                totalQtyScore = totalQtyScore + 1
            End If
        End If
    Next i
End Sub

Private Function ReceiptSetText(ByVal receiptSet As Object) As String
    Dim values As Variant, i As Long, j As Long, temp As String
    values = receiptSet.Keys
    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If CLng(Val(values(i))) > CLng(Val(values(j))) Then temp = values(i): values(i) = values(j): values(j) = temp
        Next j
    Next i
    ReceiptSetText = Join(values, "+")
End Function

Private Sub AllocateInvoiceSet(ByVal lineData As Variant, ByVal gr As Variant, ByVal invoiceId As String, ByVal vendor As String, ByVal receiptSetText As String, ByVal usedQty As Object, ByVal allocations As ListObject)
    Dim receipt As Variant
    For Each receipt In Split(receiptSetText, "+")
        AllocateInvoice lineData, gr, invoiceId, vendor, CStr(receipt), usedQty, allocations
    Next receipt
End Sub

Private Function CanAllocateExact(ByVal lineData As Variant, ByVal gr As Variant, ByVal invoiceId As String, ByVal vendor As String, ByVal receiptSetText As String, ByVal usedQty As Object, ByRef issueNote As String) As Boolean
    Dim receipts As Object, i As Long, j As Long, material As String, requiredQty As Double, availableQty As Double, sourceKey As String, receipt As Variant, issues As String, issueCount As Long
    Set receipts = CreateObject("Scripting.Dictionary")
    For Each receipt In Split(receiptSetText, "+")
        If Len(Trim$(CStr(receipt))) > 0 Then receipts(Trim$(CStr(receipt))) = True
    Next receipt
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) = invoiceId And IsDemandMaterial(CStr(lineData(i, 5))) Then
            material = CStr(lineData(i, 5))
            requiredQty = CDbl(Val(CStr(lineData(i, 7))))
            availableQty = 0
            For j = 1 To UBound(gr, 1)
                If CStr(gr(j, 4)) = vendor And CStr(gr(j, 3)) = material And receipts.Exists(CStr(gr(j, 2))) Then
                    sourceKey = CStr(gr(j, 1))
                    availableQty = availableQty + CDbl(Val(CStr(gr(j, 7))))
                    If usedQty.Exists(sourceKey) Then availableQty = availableQty - CDbl(usedQty(sourceKey))
                End If
            Next j
            If Abs(availableQty - requiredQty) >= 0.0001 Then
                If issueCount < 4 Then
                    If Len(issues) > 0 Then issues = issues & ", "
                    issues = issues & material
                End If
                issueCount = issueCount + 1
            End If
        End If
    Next i
    CanAllocateExact = (issueCount = 0)
    If Not CanAllocateExact Then
        issueNote = "QTY_NOT_EXACT: " & issues
        If issueCount > 4 Then issueNote = issueNote & " +"
    End If
End Function

Private Sub BuildReceiptIndexes(ByVal gr As Variant, ByVal receiptMaterialQty As Object, ByVal vendorReceipts As Object, ByVal receiptDates As Object, ByVal materialReceipts As Object, ByVal grRowsByBucket As Object)
    Dim i As Long, vendor As String, receipt As String, material As String, key As String, materialKey As String, qty As Double, receipts As Object, rowIndexes As Collection
    For i = 1 To UBound(gr, 1)
        vendor = CStr(gr(i, 4))
        receipt = CStr(gr(i, 2))
        material = CStr(gr(i, 3))
        qty = CDbl(Val(CStr(gr(i, 7))))
        If Len(vendor) > 0 And Len(receipt) > 0 And Len(material) > 0 And qty > 0 And IsEligibleGrRow(gr, i) Then
            mAllReceipts(receipt) = True
            key = vendor & "|" & receipt & "|" & material
            If receiptMaterialQty.Exists(key) Then
                receiptMaterialQty(key) = CDbl(receiptMaterialQty(key)) + qty
            Else
                receiptMaterialQty.Add key, qty
            End If
            materialKey = vendor & "|" & material
            If Not materialReceipts.Exists(materialKey) Then materialReceipts.Add materialKey, CreateObject("Scripting.Dictionary")
            Set receipts = materialReceipts(materialKey)
            If receipts.Exists(receipt) Then receipts(receipt) = CDbl(receipts(receipt)) + qty Else receipts.Add receipt, qty
            If Not grRowsByBucket.Exists(key) Then
                Set rowIndexes = New Collection
                grRowsByBucket.Add key, rowIndexes
            Else
                Set rowIndexes = grRowsByBucket(key)
            End If
            rowIndexes.Add i
            If Not vendorReceipts.Exists(vendor) Then vendorReceipts.Add vendor, CreateObject("Scripting.Dictionary")
            Set receipts = vendorReceipts(vendor)
            If Not receipts.Exists(receipt) Then receipts.Add receipt, True
            If Not receiptDates.Exists(vendor & "|" & receipt) And Len(CStr(gr(i, 9))) > 0 Then receiptDates.Add vendor & "|" & receipt, CDbl(Val(CStr(gr(i, 9))))
        End If
    Next i
End Sub

Private Function BuildExactAssignments(ByVal lineData As Variant, ByVal gr As Variant, ByVal invoiceId As String, ByVal vendor As String, ByVal invoiceDate As Variant, ByVal receiptDates As Object, ByVal materialReceipts As Object, ByVal grRowsByBucket As Object, ByVal usedQty As Object, ByVal preferredReceiptSet As String) As Object
    Dim assignments As Object, preferred As Object, localUsedQty As Object
    Dim i As Long, vatLineId As String, material As String, receiptItem As Variant, requiredQty As Double, receiptKey As String, materialKey As String, receipt As String, receiptsForMaterial As Object, rowIndexes As Collection
    Dim selectedReceipt As String, selectedPreferred As Boolean, candidatePreferred As Boolean
    Set assignments = CreateObject("Scripting.Dictionary")
    Set preferred = CreateObject("Scripting.Dictionary")
    Set localUsedQty = CreateObject("Scripting.Dictionary")
    For Each receiptItem In Split(preferredReceiptSet, "+")
        If Len(Trim$(CStr(receiptItem))) > 0 Then preferred(Trim$(CStr(receiptItem))) = True
    Next receiptItem
    If Not HasUsableInvoiceDate(invoiceDate) Then Set BuildExactAssignments = assignments: Exit Function
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) <> invoiceId Or Not IsDemandMaterial(CStr(lineData(i, 5))) Then GoTo NextLine
        vatLineId = CStr(lineData(i, 1))
        material = CStr(lineData(i, 5))
        requiredQty = CDbl(Val(CStr(lineData(i, 7))))
        materialKey = vendor & "|" & material
        selectedReceipt = "": selectedPreferred = False
        If materialReceipts.Exists(materialKey) Then
            Set receiptsForMaterial = materialReceipts(materialKey)
            For Each receiptItem In receiptsForMaterial.Keys
                receipt = CStr(receiptItem)
                candidatePreferred = preferred.Exists(receipt)
                If preferred.Count > 0 And Not candidatePreferred Then GoTo NextReceipt
                receiptKey = vendor & "|" & receipt & "|" & material
                If grRowsByBucket.Exists(receiptKey) Then Set rowIndexes = grRowsByBucket(receiptKey) Else GoTo NextReceipt
                If CanReserveLineFromBucket(gr, rowIndexes, requiredQty, usedQty, localUsedQty) And DateEligibilityForReceipt(invoiceDate, vendor, receipt, receiptDates) > 0 Then
                    If Len(selectedReceipt) = 0 Or (candidatePreferred And Not selectedPreferred) Or (candidatePreferred = selectedPreferred And ReceiptSortsBefore(receipt, selectedReceipt)) Then
                        selectedReceipt = receipt
                        selectedPreferred = candidatePreferred
                    End If
                End If
NextReceipt:
            Next receiptItem
        End If
        If Len(selectedReceipt) > 0 Then
            assignments(vatLineId) = selectedReceipt
            receiptKey = vendor & "|" & selectedReceipt & "|" & material
            Set rowIndexes = grRowsByBucket(receiptKey)
            ReserveLineFromBucket gr, rowIndexes, requiredQty, usedQty, localUsedQty
        End If
NextLine:
    Next i
    Set BuildExactAssignments = assignments
End Function

Private Function CanReserveLineFromBucket(ByVal gr As Variant, ByVal rowIndexes As Collection, ByVal requiredQty As Double, ByVal usedQty As Object, ByVal localUsedQty As Object) As Boolean
    Dim rowItem As Variant, rowIndex As Long, sourceKey As String, availableQty As Double, totalAvailable As Double
    For Each rowItem In rowIndexes
        rowIndex = CLng(rowItem)
        sourceKey = CStr(gr(rowIndex, 1))
        availableQty = CDbl(Val(CStr(gr(rowIndex, 7))))
        If usedQty.Exists(sourceKey) Then availableQty = availableQty - CDbl(usedQty(sourceKey))
        If localUsedQty.Exists(sourceKey) Then availableQty = availableQty - CDbl(localUsedQty(sourceKey))
        If Abs(availableQty - requiredQty) < 0.0001 Then CanReserveLineFromBucket = True: Exit Function
        If availableQty > 0 Then totalAvailable = totalAvailable + availableQty
    Next rowItem
    CanReserveLineFromBucket = (Abs(totalAvailable - requiredQty) < 0.0001)
End Function

Private Sub ReserveLineFromBucket(ByVal gr As Variant, ByVal rowIndexes As Collection, ByVal requiredQty As Double, ByVal usedQty As Object, ByVal localUsedQty As Object)
    Dim rowItem As Variant, rowIndex As Long, sourceKey As String, availableQty As Double, reserveQty As Double, exactSourceKey As String
    For Each rowItem In rowIndexes
        rowIndex = CLng(rowItem)
        sourceKey = CStr(gr(rowIndex, 1))
        availableQty = CDbl(Val(CStr(gr(rowIndex, 7))))
        If usedQty.Exists(sourceKey) Then availableQty = availableQty - CDbl(usedQty(sourceKey))
        If localUsedQty.Exists(sourceKey) Then availableQty = availableQty - CDbl(localUsedQty(sourceKey))
        If Abs(availableQty - requiredQty) < 0.0001 Then exactSourceKey = sourceKey: Exit For
    Next rowItem
    If Len(exactSourceKey) > 0 Then
        AddReservedQty localUsedQty, exactSourceKey, requiredQty
        Exit Sub
    End If
    For Each rowItem In rowIndexes
        If requiredQty <= 0 Then Exit For
        rowIndex = CLng(rowItem)
        sourceKey = CStr(gr(rowIndex, 1))
        availableQty = CDbl(Val(CStr(gr(rowIndex, 7))))
        If usedQty.Exists(sourceKey) Then availableQty = availableQty - CDbl(usedQty(sourceKey))
        If localUsedQty.Exists(sourceKey) Then availableQty = availableQty - CDbl(localUsedQty(sourceKey))
        If availableQty > 0 Then
            If availableQty < requiredQty Then reserveQty = availableQty Else reserveQty = requiredQty
            AddReservedQty localUsedQty, sourceKey, reserveQty
            requiredQty = requiredQty - reserveQty
        End If
    Next rowItem
End Sub

Private Sub AddReservedQty(ByVal localUsedQty As Object, ByVal sourceKey As String, ByVal qty As Double)
    If localUsedQty.Exists(sourceKey) Then localUsedQty(sourceKey) = CDbl(localUsedQty(sourceKey)) + qty Else localUsedQty.Add sourceKey, qty
End Sub

Private Function ReceiptSortsBefore(ByVal candidate As String, ByVal currentValue As String) As Boolean
    If IsNumeric(candidate) And IsNumeric(currentValue) Then
        ReceiptSortsBefore = (CDbl(candidate) < CDbl(currentValue))
    Else
        ReceiptSortsBefore = (StrComp(candidate, currentValue, vbTextCompare) < 0)
    End If
End Function

Private Function InvoiceMaterialDemands(ByVal lineData As Variant, ByVal invoiceId As String) As Object
    Dim demands As Object, i As Long, material As String, qty As Double
    Set demands = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) = invoiceId And IsDemandMaterial(CStr(lineData(i, 5))) Then
            material = CStr(lineData(i, 5))
            qty = CDbl(Val(CStr(lineData(i, 7))))
            If demands.Exists(material) Then demands(material) = CDbl(demands(material)) + qty Else demands.Add material, qty
        End If
    Next i
    Set InvoiceMaterialDemands = demands
End Function

Private Function InvoiceDemandLineCount(ByVal lineData As Variant, ByVal invoiceId As String) As Long
    Dim i As Long
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) = invoiceId And IsDemandMaterial(CStr(lineData(i, 5))) Then InvoiceDemandLineCount = InvoiceDemandLineCount + 1
    Next i
End Function

Private Function ReceiptSetFromAssignments(ByVal assignments As Object) As String
    Dim receipts As Object, vatLineId As Variant
    Set receipts = CreateObject("Scripting.Dictionary")
    For Each vatLineId In assignments.Keys: receipts(CStr(assignments(vatLineId))) = True: Next vatLineId
    ReceiptSetFromAssignments = ReceiptSetText(receipts)
End Function

Private Sub AllocateExactAssignments(ByVal lineData As Variant, ByVal gr As Variant, ByVal invoiceId As String, ByVal vendor As String, ByVal assignments As Object, ByVal grRowsByBucket As Object, ByVal usedQty As Object, ByVal usedReceiptMaterialQty As Object, ByVal allocations As ListObject)
    Dim i As Long, j As Long, exactRow As Long, rowItem As Variant, material As String, receipt As String, requiredQty As Double, availableQty As Double, allocationQty As Double, sourceKey As String, bucketKey As String, rowIndexes As Collection
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) = invoiceId And IsDemandMaterial(CStr(lineData(i, 5))) Then
            material = CStr(lineData(i, 5))
            If Not assignments.Exists(CStr(lineData(i, 1))) Then GoTo NextVatLine
            receipt = CStr(assignments(CStr(lineData(i, 1))))
            requiredQty = CDbl(Val(CStr(lineData(i, 7))))
            bucketKey = vendor & "|" & receipt & "|" & material
            If grRowsByBucket.Exists(bucketKey) Then
                Set rowIndexes = grRowsByBucket(bucketKey)
                exactRow = 0
                For Each rowItem In rowIndexes
                    j = CLng(rowItem)
                    sourceKey = CStr(gr(j, 1))
                    availableQty = CDbl(Val(CStr(gr(j, 7))))
                    If usedQty.Exists(sourceKey) Then availableQty = availableQty - CDbl(usedQty(sourceKey))
                    If Abs(availableQty - requiredQty) < 0.0001 Then exactRow = j: Exit For
                Next rowItem
                If exactRow > 0 Then
                    sourceKey = CStr(gr(exactRow, 1))
                    AddAllocation allocations, "A-" & CStr(allocations.ListRows.Count + 1), invoiceId, CStr(lineData(i, 1)), receipt, CLng(Val(CStr(gr(exactRow, 1)))), requiredQty, 0, "AUTO_MATCHED"
                    If usedQty.Exists(sourceKey) Then usedQty(sourceKey) = CDbl(usedQty(sourceKey)) + requiredQty Else usedQty.Add sourceKey, requiredQty
                    If usedReceiptMaterialQty.Exists(bucketKey) Then usedReceiptMaterialQty(bucketKey) = CDbl(usedReceiptMaterialQty(bucketKey)) + requiredQty Else usedReceiptMaterialQty.Add bucketKey, requiredQty
                    GoTo NextVatLine
                End If
                For Each rowItem In rowIndexes
                    j = CLng(rowItem)
                    If requiredQty > 0 Then
                    sourceKey = CStr(gr(j, 1))
                    availableQty = CDbl(Val(CStr(gr(j, 7))))
                    If usedQty.Exists(sourceKey) Then availableQty = availableQty - CDbl(usedQty(sourceKey))
                    allocationQty = WorksheetFunction.Min(requiredQty, WorksheetFunction.Max(0, availableQty))
                    If allocationQty > 0 Then
                        AddAllocation allocations, "A-" & CStr(allocations.ListRows.Count + 1), invoiceId, CStr(lineData(i, 1)), receipt, CLng(Val(CStr(gr(j, 1)))), allocationQty, requiredQty - allocationQty, "AUTO_MATCHED"
                        If usedQty.Exists(sourceKey) Then usedQty(sourceKey) = CDbl(usedQty(sourceKey)) + allocationQty Else usedQty.Add sourceKey, allocationQty
                        If usedReceiptMaterialQty.Exists(bucketKey) Then usedReceiptMaterialQty(bucketKey) = CDbl(usedReceiptMaterialQty(bucketKey)) + allocationQty Else usedReceiptMaterialQty.Add bucketKey, allocationQty
                        requiredQty = requiredQty - allocationQty
                    End If
                    End If
                Next rowItem
            End If
        End If
NextVatLine:
    Next i
End Sub

Private Function IsEligibleGrRow(ByVal gr As Variant, ByVal rowIndex As Long) As Boolean
    IsEligibleGrRow = (InStr(1, UCase$(CStr(gr(rowIndex, 10))), "QTY_DOC_ACTUAL_MISMATCH", vbTextCompare) = 0)
End Function

Public Sub EnsureMaterialVendorMap()
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("MATERIAL_NCC_MAP")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "MATERIAL_NCC_MAP"
        ws.Range("A1").Value = "Tool tu dong gan ma vat tu vao NCC theo hoa don da nhan dien MST. Khong dung cot ten NCC tu file P."
        ws.Range("A2:E2").Value = Array("MaterialNorm", "CanonicalVendor", "Source", "Active", "Note")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A2:E2"), , xlYes)
        lo.Name = "tblMaterialVendorMap"
        lo.TableStyle = "TableStyleMedium2"
    End If
End Sub

Private Sub BuildMaterialVendorIndex(ByVal invoices As ListObject, ByVal vatLines As ListObject, ByVal materialVendors As Object)
    Dim lo As ListObject, invoiceVendor As Object, i As Long, invoiceId As String, vendor As String, material As String, r As ListRow
    EnsureMaterialVendorMap
    Set lo = ThisWorkbook.Worksheets("MATERIAL_NCC_MAP").ListObjects("tblMaterialVendorMap")
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            material = Trim$(CStr(lo.DataBodyRange.Cells(i, 1).Value))
            vendor = Trim$(CStr(lo.DataBodyRange.Cells(i, 2).Value))
            If Len(material) > 0 And Len(vendor) > 0 And CBool(lo.DataBodyRange.Cells(i, 4).Value) Then materialVendors(material) = vendor
        Next i
    End If
    Set invoiceVendor = CreateObject("Scripting.Dictionary")
    For i = 1 To invoices.ListRows.Count
        invoiceId = CStr(invoices.DataBodyRange.Cells(i, 1).Value)
        vendor = Trim$(CStr(invoices.DataBodyRange.Cells(i, 7).Value))
        If Len(invoiceId) > 0 And Len(vendor) > 0 And vendor <> "UNMAPPED" Then invoiceVendor(invoiceId) = vendor
    Next i
    For i = 1 To vatLines.ListRows.Count
        invoiceId = CStr(vatLines.DataBodyRange.Cells(i, 2).Value)
        material = Trim$(CStr(vatLines.DataBodyRange.Cells(i, 5).Value))
        If invoiceVendor.Exists(invoiceId) And Len(material) > 0 And Not materialVendors.Exists(material) Then
            vendor = CStr(invoiceVendor(invoiceId))
            materialVendors.Add material, vendor
            Set r = lo.ListRows.Add
            r.Range.Cells(1, 1).Resize(1, 5).Value = Array(material, vendor, "AUTO_TU_HOA_DON", True, "Tu dong gan NCC theo hoa don da nhan dien MST.")
        End If
    Next i
End Sub

Private Sub CanonicalizeGrVendorsByMaterial(ByRef gr As Variant, ByVal materialVendors As Object)
    Dim i As Long, material As String
    For i = 1 To UBound(gr, 1)
        material = Trim$(CStr(gr(i, 3)))
        If materialVendors.Exists(material) Then
            gr(i, 4) = CStr(materialVendors(material))
        Else
            gr(i, 4) = ""
        End If
    Next i
End Sub

Private Function DateEligibilityForReceipt(ByVal invoiceDate As Variant, ByVal vendor As String, ByVal receipt As String, ByVal receiptDates As Object) As Double
    Dim key As String, dayDifference As Double, invoiceSerial As Double
    key = vendor & "|" & receipt
    If Not receiptDates.Exists(key) Then Exit Function
    If IsNumeric(invoiceDate) Then
        invoiceSerial = CDbl(invoiceDate)
    ElseIf IsDate(invoiceDate) Then
        invoiceSerial = CDbl(CDate(invoiceDate))
    Else
        Exit Function
    End If
    dayDifference = Abs(invoiceSerial - CDbl(receiptDates(key)))
    If dayDifference <= ConfigDateWindowDays() Then DateEligibilityForReceipt = 1
End Function

Private Function DatePointsForReceipt(ByVal invoiceDate As Variant, ByVal vendor As String, ByVal receipt As String, ByVal receiptDates As Object) As Double
    Dim key As String, dayDifference As Double, invoiceSerial As Double
    key = vendor & "|" & receipt
    If Not receiptDates.Exists(key) Then Exit Function
    If IsNumeric(invoiceDate) Then
        invoiceSerial = CDbl(invoiceDate)
    ElseIf IsDate(invoiceDate) Then
        invoiceSerial = CDbl(CDate(invoiceDate))
    Else
        Exit Function
    End If
    dayDifference = Abs(invoiceSerial - CDbl(receiptDates(key)))
    Select Case dayDifference
        Case 0: DatePointsForReceipt = 10
        Case 1: DatePointsForReceipt = 8
        Case 2: DatePointsForReceipt = 5
        Case Else
            If dayDifference <= ConfigDateWindowDays() Then DatePointsForReceipt = 3
    End Select
End Function

Private Function ConfigDateWindowDays() As Long
    Dim lo As ListObject, i As Long, configuredValue As Long
    configuredValue = 2
    On Error GoTo SafeExit
    Set lo = ThisWorkbook.Worksheets("CONFIG").ListObjects("tblConfig")
    If lo.DataBodyRange Is Nothing Then GoTo SafeExit
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(i, 1).Value)), "DateWindowDays", vbTextCompare) = 0 Then
            configuredValue = CLng(Val(CStr(lo.DataBodyRange.Cells(i, 2).Value)))
            Exit For
        End If
    Next i
SafeExit:
    If configuredValue < 0 Then configuredValue = 0
    ConfigDateWindowDays = configuredValue
End Function

Public Function DatePointsForTest(ByVal invoiceDate As Double, ByVal receiptDate As Double) As Double
    Dim dates As Object
    Set dates = CreateObject("Scripting.Dictionary")
    dates.Add "TEST|R", receiptDate
    DatePointsForTest = DatePointsForReceipt(invoiceDate, "TEST", "R", dates)
End Function

Public Function SearchLimitTriggered(ByVal candidateCount As Long) As Boolean
    SearchLimitTriggered = (candidateCount > MAX_RECEIPT_CANDIDATES)
End Function

Public Function MaterialScopeForTest(ByVal material As String, ByVal knownMaterial As String, ByVal externalMaterial As String) As String
    If material = knownMaterial Then
        MaterialScopeForTest = "IN_SCOPE"
    ElseIf material = externalMaterial Then
        MaterialScopeForTest = "OUT_OF_SCOPE_MATERIAL"
    Else
        MaterialScopeForTest = "UNKNOWN_MATERIAL"
    End If
End Function

Private Function EmailHintPoints(ByVal receiptSetText As String) As Double
    Dim receipt As Variant
    On Error GoTo SafeExit
    EnsureEmailHintIndex
    For Each receipt In Split(receiptSetText, "+")
        If mReceiptEmailHint.Exists(CStr(receipt)) Then EmailHintPoints = 10: Exit Function
    Next receipt
SafeExit:
End Function

Public Function EmailHintPointsForReceipt(ByVal receiptSetText As String) As Double
    EmailHintPointsForReceipt = EmailHintPoints(receiptSetText)
End Function

Private Sub EnsureEmailHintIndex()
    Dim hints As ListObject, gr As ListObject, ibs As Object, i As Long, hintData As Variant, grData As Variant
    If Not mReceiptEmailHint Is Nothing Then Exit Sub
    Set mReceiptEmailHint = CreateObject("Scripting.Dictionary")
    Set ibs = CreateObject("Scripting.Dictionary")
    Set hints = ThisWorkbook.Worksheets("EMAIL_HINTS").ListObjects("tblEmailHints")
    Set gr = ThisWorkbook.Worksheets("GR_DATA").ListObjects("tblGrData")
    If Not hints.DataBodyRange Is Nothing Then
        hintData = hints.DataBodyRange.Value2
        For i = 1 To hints.ListRows.Count
            If UCase$(CStr(hintData(i, 7))) = "HIGH" Then ibs(CStr(hintData(i, 2))) = True
        Next i
    End If
    If gr.DataBodyRange Is Nothing Then Exit Sub
    grData = gr.DataBodyRange.Value2
    For i = 1 To gr.ListRows.Count
        If ibs.Exists(CStr(grData(i, 8))) Then mReceiptEmailHint(CStr(grData(i, 2))) = True
    Next i
End Sub

Private Function CandidateReason(ByVal matchedLines As Long, ByVal lineCount As Long, ByVal datePoints As Double, ByVal hintPoints As Double) As String
    CandidateReason = "Material coverage " & CStr(matchedLines) & "/" & CStr(lineCount) & "; date=" & CStr(datePoints)
    If hintPoints > 0 Then CandidateReason = CandidateReason & "; email IB hint=10"
End Function

Private Function MergeNotes(ByVal currentNote As String, ByVal addedNote As String) As String
    If Len(currentNote) > 0 And Len(addedNote) > 0 Then MergeNotes = currentNote & " | " & addedNote Else MergeNotes = currentNote & addedNote
End Function

Private Function BuildInvoiceNote(ByVal lineData As Variant, ByVal invoiceId As String, ByVal vendor As String, ByVal receiptSet As String, ByVal materialReceipts As Object, ByVal receiptDates As Object, ByVal invoiceDate As Variant, ByVal status As String, Optional ByVal initialNote As String = "") As String
    Dim invoiceMaterials As Object, allReceipts As Object, chosenReceipts As Object, receiptsForMaterial As Object
    Dim i As Long, material As Variant, receipt As Variant, materialKey As String, noteCount As Long, notes As String, hasVendorMaterial As Boolean, hasInvoiceDate As Boolean
    notes = initialNote
    If Len(notes) > 0 Then noteCount = 1
    Set invoiceMaterials = CreateObject("Scripting.Dictionary")
    Set chosenReceipts = CreateObject("Scripting.Dictionary")
    For Each receipt In Split(receiptSet, "+")
        If Len(Trim$(CStr(receipt))) > 0 Then chosenReceipts(Trim$(CStr(receipt))) = True
    Next receipt
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) = invoiceId And IsDemandMaterial(CStr(lineData(i, 5))) Then invoiceMaterials(CStr(lineData(i, 5))) = True
    Next i
    If status = "MATCHED" Then
        BuildInvoiceNote = "Tat ca ma vat tu tren hoa don da khop voi phieu " & receiptSet & "."
        Exit Function
    End If
    hasInvoiceDate = HasUsableInvoiceDate(invoiceDate)
    If Not hasInvoiceDate Then AppendNote notes, noteCount, "Khong doc duoc ngay hoa don; chua the loc phieu theo ngay."
    If vendor = "UNMAPPED" Then AppendNote notes, noteCount, "Hoa don chua duoc nhan dien NCC theo MST, nen chua the khop phieu."
    For Each material In invoiceMaterials.Keys
        Set allReceipts = CreateObject("Scripting.Dictionary")
        hasVendorMaterial = False
        materialKey = vendor & "|" & CStr(material)
        If materialReceipts.Exists(materialKey) Then
            hasVendorMaterial = True
            Set receiptsForMaterial = materialReceipts(materialKey)
            For Each receipt In receiptsForMaterial.Keys
                If hasInvoiceDate And DateEligibilityForReceipt(invoiceDate, vendor, CStr(receipt), receiptDates) > 0 Then allReceipts(CStr(receipt)) = True
            Next receipt
        End If
        If allReceipts.Count = 0 Then
            If IsInScopeMaterial(CStr(material)) Then
                If hasVendorMaterial And hasInvoiceDate Then
                    AppendNote notes, noteCount, "Ma " & CStr(material) & " khong tim thay phieu trong khoang 2 ngay cua hoa don."
                ElseIf Not hasVendorMaterial Then
                    AppendNote notes, noteCount, "Ma " & CStr(material) & " co trong file P nhung chua duoc cau hinh ve NCC " & vendor & "."
                End If
            Else
                AppendNote notes, noteCount, "Ma " & CStr(material) & " co tren hoa don nhung khong co trong file P."
            End If
        ElseIf allReceipts.Count >= 3 Then
            AppendNote notes, noteCount, "Ma " & CStr(material) & " tren hoa don tim thay o " & CStr(allReceipts.Count) & " phieu (vi du: " & ReceiptExamples(allReceipts) & "); chua du co so de khop."
        End If
    Next material
    If Len(notes) = 0 Then notes = "Chua tim thay phieu co tat ca ma vat tu khop voi hoa don."
    BuildInvoiceNote = notes
End Function

Private Function HasUsableInvoiceDate(ByVal invoiceDate As Variant) As Boolean
    If IsNumeric(invoiceDate) Then HasUsableInvoiceDate = (CDbl(invoiceDate) > 0) Else HasUsableInvoiceDate = IsDate(invoiceDate)
End Function

Private Function IsReceiptWithinDate(ByVal invoiceDate As Variant, ByVal receiptDate As Variant) As Boolean
    Dim invoiceSerial As Double, receiptSerial As Double
    If Not HasUsableInvoiceDate(invoiceDate) Then Exit Function
    If IsNumeric(invoiceDate) Then invoiceSerial = CDbl(invoiceDate) Else invoiceSerial = CDbl(CDate(invoiceDate))
    If IsNumeric(receiptDate) Then
        receiptSerial = CDbl(receiptDate)
    ElseIf IsDate(receiptDate) Then
        receiptSerial = CDbl(CDate(receiptDate))
    Else
        Exit Function
    End If
    IsReceiptWithinDate = (Abs(invoiceSerial - receiptSerial) <= ConfigDateWindowDays())
End Function

Private Function ReceiptExamples(ByVal receipts As Object) As String
    Dim keys As Variant, i As Long, result As String, maxItems As Long
    keys = receipts.Keys
    maxItems = WorksheetFunction.Min(5, receipts.Count)
    For i = LBound(keys) To LBound(keys) + maxItems - 1
        If Len(result) > 0 Then result = result & ","
        result = result & CStr(keys(i))
    Next i
    ReceiptExamples = result
End Function

Private Function MaterialExistsInP(ByVal gr As Variant, ByVal material As String) As Boolean
    Dim i As Long
    For i = 1 To UBound(gr, 1)
        If CStr(gr(i, 3)) = material Then MaterialExistsInP = True: Exit Function
    Next i
End Function

Private Function AnyReceiptInSet(ByVal allReceipts As Object, ByVal chosenReceipts As Object) As Boolean
    Dim receipt As Variant
    For Each receipt In allReceipts.Keys
        If chosenReceipts.Exists(CStr(receipt)) Then AnyReceiptInSet = True: Exit Function
    Next receipt
End Function

Private Sub AppendNote(ByRef notes As String, ByRef noteCount As Long, ByVal item As String)
    If noteCount >= 4 Then
        If InStr(1, notes, "Con ma nghi ngo khac", vbTextCompare) = 0 Then notes = notes & " | Con ma nghi ngo khac; xem VAT_LINES."
        Exit Sub
    End If
    If Len(notes) > 0 Then notes = notes & " | "
    notes = notes & item
    noteCount = noteCount + 1
End Sub

Private Sub ScoreReceipt(ByVal lineData As Variant, ByVal invoiceId As String, ByVal vendor As String, ByVal receipt As String, ByVal receiptMaterialQty As Object, ByRef lineCount As Long, ByRef matchedLines As Long, ByRef totalQtyScore As Double)
    Dim i As Long, material As String, requiredQty As Double, availableQty As Double, key As String
    lineCount = 0: matchedLines = 0: totalQtyScore = 0
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) = invoiceId And IsDemandMaterial(CStr(lineData(i, 5))) Then
            lineCount = lineCount + 1
            material = CStr(lineData(i, 5))
            requiredQty = CDbl(Val(CStr(lineData(i, 7))))
            key = vendor & "|" & receipt & "|" & material
            If receiptMaterialQty.Exists(key) Then
                availableQty = CDbl(receiptMaterialQty(key))
            If Abs(availableQty - requiredQty) < 0.0001 Then
                matchedLines = matchedLines + 1
                totalQtyScore = totalQtyScore + 1
            End If
            End If
        End If
    Next i
End Sub

Private Sub AllocateInvoice(ByVal lineData As Variant, ByVal gr As Variant, ByVal invoiceId As String, ByVal vendor As String, ByVal receipt As String, ByVal usedQty As Object, ByVal allocations As ListObject)
    Dim i As Long, j As Long, material As String, requiredQty As Double, availableQty As Double, allocationQty As Double, key As String
    For i = 1 To UBound(lineData, 1)
        If CStr(lineData(i, 2)) = invoiceId And IsDemandMaterial(CStr(lineData(i, 5))) Then
            material = CStr(lineData(i, 5))
            requiredQty = CDbl(Val(CStr(lineData(i, 7))))
            For j = 1 To UBound(gr, 1)
                If CStr(gr(j, 4)) = vendor And CStr(gr(j, 2)) = receipt And CStr(gr(j, 3)) = material And requiredQty > 0 Then
                    key = CStr(gr(j, 1))
                    availableQty = CDbl(Val(CStr(gr(j, 7))))
                    If usedQty.Exists(key) Then availableQty = availableQty - CDbl(usedQty(key))
                    allocationQty = WorksheetFunction.Min(requiredQty, WorksheetFunction.Max(0, availableQty))
                    If allocationQty > 0 Then
                        AddAllocation allocations, "A-" & CStr(allocations.ListRows.Count + 1), invoiceId, CStr(lineData(i, 1)), receipt, CLng(Val(CStr(gr(j, 1)))), allocationQty, requiredQty - allocationQty, "AUTO_MATCHED"
                        If usedQty.Exists(key) Then usedQty(key) = CDbl(usedQty(key)) + allocationQty Else usedQty.Add key, allocationQty
                        requiredQty = requiredQty - allocationQty
                    End If
                End If
            Next j
        End If
    Next i
End Sub

Private Sub AddCandidate(ByVal lo As ListObject, ByVal candidateId As String, ByVal invoiceId As String, ByVal receiptSet As String, ByVal vendorScore As Double, ByVal materialScore As Double, ByVal qtyScore As Double, ByVal dateScore As Double, ByVal hintScore As Double, ByVal totalScore As Double, ByVal rank As Long, ByVal reasons As String)
    Dim r As ListRow
    Set r = lo.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 11).Value = Array(candidateId, invoiceId, receiptSet, vendorScore, materialScore, qtyScore, dateScore, hintScore, totalScore, rank, reasons)
End Sub

Private Sub RankCandidatesForInvoice(ByVal lo As ListObject, ByVal invoiceId As String)
    Dim i As Long, j As Long, rankValue As Long, scoreValue As Double, otherScore As Double
    Dim receiptValue As String, otherReceipt As String
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, 2).Value) = invoiceId Then
            scoreValue = CDbl(Val(CStr(lo.DataBodyRange.Cells(i, 9).Value)))
            receiptValue = CStr(lo.DataBodyRange.Cells(i, 3).Value)
            rankValue = 1
            For j = 1 To lo.ListRows.Count
                If j <> i And CStr(lo.DataBodyRange.Cells(j, 2).Value) = invoiceId Then
                    otherScore = CDbl(Val(CStr(lo.DataBodyRange.Cells(j, 9).Value)))
                    otherReceipt = CStr(lo.DataBodyRange.Cells(j, 3).Value)
                    If otherScore > scoreValue Or (Abs(otherScore - scoreValue) < 0.000001 And ReceiptSortsBefore(otherReceipt, receiptValue)) Then rankValue = rankValue + 1
                End If
            Next j
            lo.DataBodyRange.Cells(i, 10).Value = rankValue
        End If
    Next i
End Sub

Private Sub AddAllocation(ByVal lo As ListObject, ByVal allocationId As String, ByVal invoiceId As String, ByVal vatLineId As String, ByVal receiptNo As String, ByVal sourceRow As Long, ByVal allocatedQty As Double, ByVal residual As Double, ByVal status As String)
    Dim r As ListRow
    Set r = lo.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 8).Value = Array(allocationId, invoiceId, vatLineId, receiptNo, sourceRow, allocatedQty, residual, status)
End Sub

Private Sub AddInvoiceReport(ByVal lo As ListObject, ByVal runId As String, ByVal invData As Variant, ByVal index As Long, ByVal receipts As String, ByVal score As Double, ByVal status As String, Optional ByVal reasonOverride As String = "")
    Dim r As ListRow, reason As String
    If status = "MATCHED" Then reason = Viet("0054_1EA5_0074_0020_0063_1EA3_0020_006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_0074_0072_00EA_006E_0020_0068_00F3_0061_0020_0111_01A1_006E_0020_0111_00E3_0020_0063_00F3_0020_0074_0072_006F_006E_0067_0020_0070_0068_0069_1EBF_0075_002E") Else reason = Viet("0043_0068_01B0_0061_0020_0111_1EE7_0020_006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_0111_1EC3_0020_006B_0068_1EDB_0070_002E_0020_0043_1EA7_006E_0020_006B_0069_1EC3_006D_0020_0074_0072_0061_0020_006C_1EA1_0069_002E")
    If Len(reasonOverride) > 0 Then reason = reasonOverride
    reason = VietHoaLyDo(reason)
    Set r = lo.ListRows.Add
    r.Range.Cells(1, 1).Resize(1, 11).Value = Array(runId, CStr(invData(index, 2)), "", CStr(invData(index, 7)), invData(index, 6), receipts, score, status, "", "", reason)
    r.Range.Cells(1, 3).NumberFormat = "@"
    r.Range.Cells(1, 3).Value = CStr(invData(index, 5))
End Sub

Private Function VietHoaLyDo(ByVal reason As String) As String
    reason = Replace(reason, "Tat ca ma vat tu tren hoa don da khop voi phieu ", Viet("0054_1EA5_0074_0020_0063_1EA3_0020_006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_0074_0072_00EA_006E_0020_0068_00F3_0061_0020_0111_01A1_006E_0020_0111_00E3_0020_006B_0068_1EDB_0070_0020_0076_1EDB_0069_0020_0070_0068_0069_1EBF_0075_0020"))
    reason = Replace(reason, " co tren hoa don nhung khong co trong file P.", Viet("0020_0063_00F3_0020_0074_0072_00EA_006E_0020_0068_00F3_0061_0020_0111_01A1_006E_0020_006E_0068_01B0_006E_0067_0020_006B_0068_00F4_006E_0067_0020_0063_00F3_0020_0074_0072_006F_006E_0067_0020_0066_0069_006C_0065_0020_0050_002E"))
    reason = Replace(reason, " tren hoa don tim thay o ", Viet("0020_0074_0072_00EA_006E_0020_0068_00F3_0061_0020_0111_01A1_006E_0020_0074_00EC_006D_0020_0074_0068_1EA5_0079_0020_1EDF_0020"))
    reason = Replace(reason, " phieu (vi du: ", Viet("0020_0070_0068_0069_1EBF_0075_0020_0028_0076_00ED_0020_0064_1EE5_003A_0020"))
    reason = Replace(reason, "); chua du co so de khop.", Viet("0029_002C_0020_0063_0068_01B0_0061_0020_0111_1EE7_0020_0063_01A1_0020_0073_1EDF_0020_0111_1EC3_0020_006B_0068_1EDB_0070_002E"))
    reason = Replace(reason, "Khong doc duoc ngay hoa don; chua the loc phieu theo ngay.", Viet("004B_0068_00F4_006E_0067_0020_0111_1ECD_0063_0020_0111_01B0_1EE3_0063_0020_006E_0067_00E0_0079_0020_0068_00F3_0061_0020_0111_01A1_006E_003B_0020_0063_0068_01B0_0061_0020_0074_0068_1EC3_0020_006C_1ECD_0063_0020_0070_0068_0069_1EBF_0075_0020_0074_0068_0065_006F_0020_006E_0067_00E0_0079_002E"))
    reason = Replace(reason, " khong tim thay phieu trong khoang 2 ngay cua hoa don.", Viet("0020_006B_0068_00F4_006E_0067_0020_0074_00EC_006D_0020_0074_0068_1EA5_0079_0020_0070_0068_0069_1EBF_0075_0020_0074_0072_006F_006E_0067_0020_006B_0068_006F_1EA3_006E_0067_0020_0032_0020_006E_0067_00E0_0079_0020_0063_1EE7_0061_0020_0068_00F3_0061_0020_0111_01A1_006E_002E"))
    reason = Replace(reason, "Hoa don chua duoc nhan dien NCC theo MST, nen chua the khop phieu.", Viet("0048_00F3_0061_0020_0111_01A1_006E_0020_0063_0068_01B0_0061_0020_0111_01B0_1EE3_0063_0020_006E_0068_1EAD_006E_0020_0064_0069_1EC7_006E_0020_004E_0043_0043_0020_0074_0068_0065_006F_0020_004D_0053_0054_002C_0020_006E_00EA_006E_0020_0063_0068_01B0_0061_0020_0074_0068_1EC3_0020_006B_0068_1EDB_0070_0020_0070_0068_0069_1EBF_0075_002E"))
    reason = Replace(reason, " co trong file P nhung chua duoc cau hinh ve NCC ", Viet("0020_0063_00F3_0020_0074_0072_006F_006E_0067_0020_0066_0069_006C_0065_0020_0050_0020_006E_0068_01B0_006E_0067_0020_0063_0068_01B0_0061_0020_0111_01B0_1EE3_0063_0020_0063_1EA5_0075_0020_0068_00EC_006E_0068_0020_0076_1EC1_0020_004E_0043_0043_0020"))
    reason = Replace(reason, " khong du so luong tren phieu de xuat.", Viet("0020_006B_0068_00F4_006E_0067_0020_0111_1EE7_0020_0073_1ED1_0020_006C_01B0_1EE3_006E_0067_0020_0074_0072_00EA_006E_0020_0070_0068_0069_1EBF_0075_0020_0111_1EC1_0020_0078_0075_1EA5_0074_002E"))
    reason = Replace(reason, "Con ma nghi ngo khac; xem VAT_LINES.", Viet("0043_00F2_006E_0020_006D_00E3_0020_006E_0067_0068_0069_0020_006E_0067_1EDD_0020_006B_0068_00E1_0063_003B_0020_0078_0065_006D_0020_0056_0041_0054_005F_004C_0049_004E_0045_0053_002E"))
    reason = Replace(reason, "Con ma vat tu khac khong du so luong; xem VAT_LINES.", Viet("0043_00F2_006E_0020_006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_006B_0068_00E1_0063_0020_006B_0068_00F4_006E_0067_0020_0111_1EE7_0020_0073_1ED1_0020_006C_01B0_1EE3_006E_0067_003B_0020_0078_0065_006D_0020_0056_0041_0054_005F_004C_0049_004E_0045_0053_002E"))
    reason = Replace(reason, "QTY_NOT_EXACT: ", Viet("0053_1ED1_0020_006C_01B0_1EE3_006E_0067_0020_0074_0072_00EA_006E_0020_0068_00F3_0061_0020_0111_01A1_006E_0020_006B_0068_00F4_006E_0067_0020_0062_1EB1_006E_0067_0020_0074_1ED5_006E_0067_0020_0073_1ED1_0020_006C_01B0_1EE3_006E_0067_0020_1EDF_0020_0070_0068_0069_1EBF_0075_0020_0111_1EC1_0020_0078_0075_1EA5_0074_003A_0020"))
    reason = Replace(reason, "Ma ", Viet("004D_00E3_0020"))
    reason = Replace(reason, "All invoice materials covered by receipt.", Viet("0054_1EA5_0074_0020_0063_1EA3_0020_006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_0074_0072_00EA_006E_0020_0068_00F3_0061_0020_0111_01A1_006E_0020_0111_00E3_0020_0063_00F3_0020_0074_0072_006F_006E_0067_0020_0070_0068_0069_1EBF_0075_002E"))
    reason = Replace(reason, "Tat ca ma vat tu tren hoa don da co trong phieu.", Viet("0054_1EA5_0074_0020_0063_1EA3_0020_006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_0074_0072_00EA_006E_0020_0068_00F3_0061_0020_0111_01A1_006E_0020_0111_00E3_0020_0063_00F3_0020_0074_0072_006F_006E_0067_0020_0070_0068_0069_1EBF_0075_002E"))
    reason = Replace(reason, "Incomplete material coverage; manual review required.", Viet("0043_0068_01B0_0061_0020_0111_1EE7_0020_006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_0111_1EC3_0020_006B_0068_1EDB_0070_002E_0020_0043_1EA7_006E_0020_006B_0069_1EC3_006D_0020_0074_0072_0061_0020_006C_1EA1_0069_002E"))
    reason = Replace(reason, "Chua du ma vat tu de khop. Can kiem tra lai.", Viet("0043_0068_01B0_0061_0020_0111_1EE7_0020_006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_0111_1EC3_0020_006B_0068_1EDB_0070_002E_0020_0043_1EA7_006E_0020_006B_0069_1EC3_006D_0020_0074_0072_0061_0020_006C_1EA1_0069_002E"))
    reason = Replace(reason, "NO_IN_SCOPE_MATERIAL:", Viet("004B_0068_00F4_006E_0067_0020_0063_00F3_0020_006D_00E3_0020_0074_0072_006F_006E_0067_0020_0070_0068_0069_1EBF_0075_003A"))
    reason = Replace(reason, "all ", Viet("0054_1EA5_0074_0020_0063_1EA3_0020"))
    reason = Replace(reason, "VAT line(s) are outside the GR snapshot.", Viet("006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_006E_1EB1_006D_0020_006E_0067_006F_00E0_0069_0020_0064_1EEF_0020_006C_0069_1EC7_0075_0020_0047_0052_002E"))
    reason = Replace(reason, "SEARCH_TRUNCATED:", Viet("0054_00EC_006D_0020_006B_0069_1EBF_006D_0020_0071_0075_00E1_0020_006E_0068_0069_1EC1_0075_003A"))
    reason = Replace(reason, "more than ", Viet("006E_0068_0069_1EC1_0075_0020_0068_01A1_006E_0020"))
    reason = Replace(reason, "receipt candidates; refine vendor/date filters.", Viet("0070_0068_0069_1EBF_0075_0020_0111_1EC1_0020_0078_0075_1EA5_0074_003B_0020_0068_00E3_0079_0020_0074_0068_0075_0020_0068_1EB9_0070_0020_004E_0043_0043_002F_006E_0067_00E0_0079_002E"))
    reason = Replace(reason, "MIXED_SCOPE:", Viet("004C_1EAB_006E_0020_006D_00E3_0020_006E_0067_006F_00E0_0069_0020_0078_01B0_1EDF_006E_0067_003A"))
    reason = Replace(reason, "out-of-scope VAT line(s) ignored for allocation.", Viet("006D_00E3_0020_0076_1EAD_0074_0020_0074_01B0_0020_006E_0067_006F_00E0_0069_0020_0078_01B0_1EDF_006E_0067_0020_006B_0068_00F4_006E_0067_0020_0074_00ED_006E_0068_0020_006B_0068_0069_0020_0070_0068_00E2_006E_0020_0062_1ED5_002E"))
    reason = Replace(reason, "CAPACITY_CONFLICT:", Viet("0056_01B0_1EE3_0074_0020_0073_1ED1_0020_006C_01B0_1EE3_006E_0067_0020_0070_0068_0069_1EBF_0075_003A"))
    reason = Replace(reason, "allocated ", Viet("0111_00E3_0020_0070_0068_00E2_006E_0020_0062_1ED5_0020"))
    reason = Replace(reason, " of ", "/")
    VietHoaLyDo = reason
End Function

Private Function Viet(ByVal codePoints As String) As String
    Dim parts() As String, i As Long
    parts = Split(codePoints, "_")
    For i = LBound(parts) To UBound(parts): Viet = Viet & ChrW$(CLng("&H" & parts(i))): Next i
End Function

Private Sub ClearRows(ByVal lo As ListObject)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
End Sub
