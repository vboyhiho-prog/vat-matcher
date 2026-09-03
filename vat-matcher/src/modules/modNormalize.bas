Attribute VB_Name = "modNormalize"
Option Explicit

Public Function NormalizeText(ByVal value As String) As String
    Dim s As String
    s = UCase$(Trim$(Replace(Replace(value, vbCr, " "), vbLf, " ")))
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeText = StripVietnameseMarks(s)
End Function

Public Function NormalizeMaterial(ByVal value As String) As String
    Dim s As String, ch As String, i As Long
    s = UCase$(Trim$(value))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Z0-9]" Then NormalizeMaterial = NormalizeMaterial & ch
    Next i
End Function

Public Function NormalizeVendor(ByVal value As String) As String
    Dim s As String, lo As ListObject, i As Long, aliasValue As String, fileAlias As String
    s = NormalizeText(value)
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("NCC_MAP").ListObjects("tblVendorMap")
    On Error GoTo 0
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            aliasValue = NormalizeText(CStr(lo.DataBodyRange.Cells(i, 3).Value))
            fileAlias = NormalizeText(CStr(lo.DataBodyRange.Cells(i, 4).Value))
            If Len(aliasValue) > 0 And (s = aliasValue Or InStr(1, s, aliasValue, vbTextCompare) > 0) Then NormalizeVendor = CStr(lo.DataBodyRange.Cells(i, 1).Value): Exit Function
            If Len(fileAlias) > 0 And (s = fileAlias Or InStr(1, s, fileAlias, vbTextCompare) > 0) Then NormalizeVendor = CStr(lo.DataBodyRange.Cells(i, 1).Value): Exit Function
        Next i
    End If
    If InStr(s, "THANH DAT") > 0 Then NormalizeVendor = "THANH_DAT": Exit Function
    If InStr(s, "LTV") > 0 Or InStr(s, "CONG NGHIEP LTV") > 0 Then NormalizeVendor = "LTV": Exit Function
    NormalizeVendor = s
End Function

Private Function StripVietnameseMarks(ByVal s As String) As String
    ReplaceCodes s, Array(&HC0, &HC1, &HC2, &HC3, &H102, &H1EA0, &H1EA2, &H1EA4, &H1EA6, &H1EA8, &H1EAA, &H1EAC, &H1EB0, &H1EB2, &H1EB4, &H1EB6), "A"
    ReplaceCodes s, Array(&HC8, &HC9, &HCA, &H1EBC, &H1EB8, &H1EBA, &H1EBC, &H1EBE, &H1EC0, &H1EC2, &H1EC4, &H1EC6), "E"
    ReplaceCodes s, Array(&HCC, &HCD, &H128, &H1EC8, &H1ECA), "I"
    ReplaceCodes s, Array(&HD2, &HD3, &HD4, &HD5, &H1ECC, &H1ECE, &H1ED0, &H1ED2, &H1ED4, &H1ED6, &H1ED8, &H1EDA, &H1EDC, &H1EDE, &H1EE0, &H1EE2, &H1A0), "O"
    ReplaceCodes s, Array(&HD9, &HDA, &H168, &H1EE4, &H1EE6, &H1EE8, &H1EEA, &H1EEC, &H1EEE, &H1EF0, &H1AF), "U"
    ReplaceCodes s, Array(&HDD, &H1EF2, &H1EF4, &H1EF6, &H1EF8), "Y"
    s = Replace(s, ChrW(&H110), "D")
    StripVietnameseMarks = s
End Function

Private Sub ReplaceCodes(ByRef s As String, ByVal codes As Variant, ByVal replacement As String)
    Dim i As Long
    For i = LBound(codes) To UBound(codes): s = Replace(s, ChrW(CLng(codes(i))), replacement): Next i
End Sub

Public Function QtyValue(ByVal value As Variant) As Variant
    If IsNumeric(value) Then QtyValue = CDbl(value) Else QtyValue = Empty
End Function
