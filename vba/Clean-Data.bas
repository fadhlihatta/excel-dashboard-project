Attribute VB_Name = "Module2"
Sub CleanData()

Dim rdsheet As Worksheet
Dim cdsheet As Worksheet
Dim rdlastrow As Long
Dim cdlastrow As Long
Dim lastcol As Long
Dim rng As Range
Dim arrdata As Variant
Dim i As Long
Dim tbl As ListObject

Set rdsheet = Sheets("RawData")
Set cdsheet = Sheets("CleanedData")

rdlastrow = rdsheet.Cells(Rows.Count, 1).End(xlUp).Row
cdlastrow = cdsheet.Cells(Rows.Count, 1).End(xlUp).Row

If rdlastrow <= cdlastrow Then
    MsgBox "No new Data!"
    Exit Sub
End If

lastcol = rdsheet.Cells(1, Columns.Count).End(xlToLeft).Column

Set rng = rdsheet.Range(rdsheet.Cells(cdlastrow, 1), rdsheet.Cells(rdlastrow, lastcol))

arrdata = rng.Value

For i = 1 To UBound(arrdata, 1)

    If IsEmpty(arrdata(i, 1)) Or IsError(arrdata(i, 1)) Then GoTo nextcell
    
    arrdata(i, 1) = Trim(arrdata(i, 1))
    
    If IsDate(arrdata(i, 2)) Then
        arrdata(i, 2) = CDate(arrdata(i, 2))
    ElseIf IsNumeric(arrdata(i, 1)) Then
            arrdata(i, 2) = CDbl(arrdata(i, 1))
    End If
    
nextcell:
Next i

cdsheet.Range(cdsheet.Cells(cdlastrow, 1), cdsheet.Cells(rdlastrow, lastcol)).Value = arrdata

If cdsheet.ListObjects.Count > 0 Then
    Set tbl = cdsheet.ListObjects(1)
Else
    Set tbl = cdsheet.ListObjects.Add(xlSrcRange, cdsheet.Range("A1").CurrentRegion, , xlYes)
End If

tbl.DataBodyRange.RemoveDuplicates

End Sub
