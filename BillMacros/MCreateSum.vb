Module MCreateSum
    Sub CreateSum()
        'This sub creates a summary based on the billsheets
        'A summary sheet is inserted if it does not exist
        Dim BillSheets As Excel.Sheets
        Dim Wksht As Excel.Worksheet, SumSheet As Excel.Worksheet
        Dim SumRow As Integer
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        BillSheets = XlWb.Worksheets
        ShowActivationNotice() 'Show activation warning window

        SumSheet = GetSumSheet() 'Insert SumSheet if it does not exist

        With SumSheet
            LogTrackInfo("CreateSum")
            If .Columns(1).Find("BillGrpStart") Is Nothing Then
                MsgBox("Cannot find BillGrpStart", vbCritical, "Excel Bill Functions")
                Exit Sub
            ElseIf .Columns(1).Find("BillGrpEnd") Is Nothing Then
                MsgBox("Cannot find BillGrpEnd", vbCritical, "Excel Bill Functions")
                Exit Sub
            Else 'Bill group exits
                If Not NamedRangeExists(GetSumTemplateSheet, "SumBillRow") Then
                    MsgBox("Cannot find named range SumBillRow", vbCritical, "Excel Bill Functions")
                    Exit Sub
                Else 'SumBillRow exists
                    xlAp.ScreenUpdating = False
                    'Delete all rows in billgroup
                    SumRow = .Columns(1).Find("BillGrpStart").Row + 1
                    Do Until .Cells(SumRow, 1).value = "BillGrpEnd"
                        .Rows(SumRow).Delete
                    Loop
                    'Populate bill group
                    For Each Wksht In BillSheets
                        If Wksht.Cells(1, 1).value = "#BillSheet#" And Wksht.Tab.Color = RGB(255, 0, 0) Then
                            .Rows(SumRow).Insert(shift:=Excel.XlDirection.xlDown)
                            InsertSumRow(SumSheet, Wksht, SumRow)
                            .Rows(SumRow).AutoFit
                            SumRow = SumRow + 1
                        End If
                    Next
                End If 'NamedRangeExists
            End If
            SumSheet.Activate()
        End With
        xlAp.ScreenUpdating = True
    End Sub
    Sub InsertSumRow(SumSheet As Excel.Worksheet, Wksht As Excel.Worksheet, SumRow As Integer)
        'Replace sheet names in each formula cell of SumBillRow with the WkSht name and
        'Insert the SumBillRow range from #SumTemplate#
        'Adjust column widths according to SumBillRow
        Dim SumTemplate As Excel.Worksheet
        Dim Cell As Excel.Range
        Dim SumBillRowCols As Integer
        Dim SumBillRowRows As Integer
        Dim NewSumBillRow As Excel.Range
        SumTemplate = GetSumTemplateSheet()
        SumBillRowCols = SumTemplate.Range("SumBillRow").Columns.Count
        SumBillRowRows = SumTemplate.Range("SumBillRow").Rows.Count
        SumTemplate.Range("SumBillRow").Copy(SumSheet.Cells(SumRow, 1))
        NewSumBillRow = SumSheet.Range(SumSheet.Cells(SumRow, 1), SumSheet.Cells(SumRow, 1).Offset(SumBillRowRows - 1, SumBillRowCols - 1))
        For Each Cell In NewSumBillRow
            If Cell.HasFormula Then
                Cell.Formula = ReplaceFormulaRefs(Cell.Formula, "'" & Wksht.Name & "'!")
            End If
        Next
        NewSumBillRow.Copy(Destination:=SumSheet.Cells(SumRow, 1))
    End Sub


End Module
