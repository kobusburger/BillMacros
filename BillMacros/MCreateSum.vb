Module MCreateSum
    Sub CreateSum()
        'This sub creates a summary based on the billsheets and BillSheetTemplate sheet
        'A summary sheet is inserted if it does not exist
        Dim Wksht As Excel.Worksheet, SumSheet As Excel.Worksheet
        Dim SumRow As Integer
        Dim BillSheets As Excel.Sheets
        xlWb = xlAp.ActiveWorkbook
        xlSh = xlWb.ActiveSheet
        BillSheets = xlWb.Worksheets
        ShowActivationNotice() 'Show termination warning

        SumSheet = GetSumSheet() 'Insert SumSheet if it does not exist
        CheckTemplateSheet("SumTemplate") 'Check SumTemplate sheet and named ranges and insert/ replace if not correct

        LogTrackInfo("CreateSum")
        With SumSheet
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
                    SumRow += 1
                End If
            Next
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

        SumTemplate = XlWb.Worksheets("SumTemplate")
        SumBillRowCols = SumTemplate.Range("SumBillRow").Columns.Count
        SumBillRowRows = SumTemplate.Range("SumBillRow").Rows.Count
        SumTemplate.Range("SumBillRow").Copy(SumSheet.Cells(SumRow, 1))
        NewSumBillRow = SumSheet.Range(SumSheet.Cells(SumRow, 1), SumSheet.Cells(SumRow, 1).Offset(SumBillRowRows - 1, SumBillRowCols - 1))
        For Each Cell In NewSumBillRow
            If Cell.HasFormula Then
                If TypeOf Cell.Value IsNot Int32 Then 'Only replace formulas in error free cells
                    Cell.Formula = ReplaceFormulaRefs(Cell.Formula, "'" & Wksht.Name & "'!")
                End If
            End If
        Next
        NewSumBillRow.Copy(Destination:=SumSheet.Cells(SumRow, 1))
    End Sub
    Function GetSumSheet() As Excel.Worksheet
        'Search for "Summary" sheet and insert if it does not exist or if it is not correct
        '        Dim Wksht As Excel.Worksheet

        Dim BillSheets As Excel.Sheets
        BillSheets = xlWb.Worksheets
        On Error Resume Next
        GetSumSheet = BillSheets("Summary")
        On Error GoTo 0
        If GetSumSheet Is Nothing Then
            CreateSheet("Summary", Excel.XlRgbColor.rgbGreen, True) 'It does not exist
        ElseIf GetSumSheet.Columns(1).Find("BillGrpStart") Is Nothing Or
                GetSumSheet.Columns(1).Find("BillGrpEnd") Is Nothing Then
            CreateSheet("Summary", Excel.XlRgbColor.rgbGreen, True) 'There is a problem with it
        End If
        GetSumSheet = BillSheets("Summary")
        GetSumSheet.Tab.Color = Excel.XlRgbColor.rgbGreen 'Confirm the tab color
    End Function
End Module
