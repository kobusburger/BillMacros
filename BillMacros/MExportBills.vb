Module MExportBills
    Sub CreatePDF()
        'Save-as blank PDF
        'Only works in Excel 2007 and later

        Dim Wksht As Excel.Worksheet, StartSht As Excel.Worksheet
        Dim result As Boolean, First As Boolean
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet

        StartSht = XlSh
        ShowActivationNotice() 'Show activation warning window
        First = True
        For Each Wksht In XlWb.Worksheets
            Select Case Wksht.Tab.Color
                Case Drawing.Color.Red 'Red = BillSheet
                    If First Then
                        Wksht.Select(True)
                        First = False
                    Else
                        Wksht.Select(False)
                    End If
                Case Drawing.Color.Green 'Green = Summary
                    Wksht.Select(False)
            End Select
        Next
        result = xlAp.Dialogs.xlDialogSaveAs.Show(, 57) 'pdf type_num = 57
        StartSht.Select()
    End Sub
    Sub CreateStripped()
        'Delete hidden rows, delete non-bill columns & delete non-bill sheets
        Dim Wksht As Excel.Worksheet, FName As String
        'Save bill with new name
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        FName = Left(XlWb.Name, (InStrRev(XlWb.Name, ".", -1, vbTextCompare) - 1))
        If Not xlAp.Dialogs(Excel.XlBuiltInDialog.xlDialogSaveAs).Show(FName & " Stripped") Then Exit Sub

        For Each Wksht In XlWb.Worksheets
            Wksht.Visible = Excel.XlSheetVisibility.xlSheetVisible 'Worksheets must be visible to avoid errors
            Select Case Wksht.Tab.Color
                Case Drawing.Color.Red 'Red = BillSheet
                    Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas
                    DeleteXtraRowsCols(Wksht, "#BillEnd#", AmtCol)
                Case Drawing.Color.Green 'Green = Summary
                    Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas
                    DeleteXtraRowsCols(Wksht, "#SumEnd#", SumAmtCol)
                Case Else
                    xlAp.DisplayAlerts = False
                    Wksht.Delete()
                    xlAp.DisplayAlerts = True
            End Select
        Next

    End Sub
    Sub CreatePriced()
        'Delete hidden rows, delete non-bill columns, delete non-bill sheets & copy priced columns to bill
        Dim Wksht As Excel.Worksheet, FName As String
        Dim MaxRowNo As Integer, MaxColNo As Integer
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        'Save bill with new name
        FName = Left(XlWb.Name, (InStrRev(XlWb.Name, ".", -1, vbTextCompare) - 1))
        If Not xlAp.Dialogs(Excel.XlBuiltInDialog.xlDialogSaveAs).Show(FName & " Priced") Then Exit Sub

        MaxRowNo = xlAp.Rows.Count
        MaxColNo = xlAp.Columns.Count
        For Each Wksht In XlWb.Worksheets
            Wksht.Visible = Excel.XlSheetVisibility.xlSheetVisible 'Worksheets must be visible to avoid errors
            Select Case Wksht.Tab.Color
                Case Drawing.Color.Red 'Red = BillSheet
                    Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas
                    Wksht.Columns(PricedAmtCol).Copy(Wksht.Columns(AmtCol))
                    Wksht.Columns(PricedRateCol).Copy(Wksht.Columns(RateCol))
                    DeleteXtraRowsCols(Wksht, "#BillEnd#", AmtCol)
                Case Drawing.Color.Green 'Green = Summary
                    Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas
                    Wksht.Columns(SumPricedAmtCol).Copy(Wksht.Columns(SumAmtCol))
                    DeleteXtraRowsCols(Wksht, "#SumEnd#", SumAmtCol)
                Case Else
                    xlAp.DisplayAlerts = False
                    Wksht.Delete()
                    xlAp.DisplayAlerts = True
            End Select
        Next

    End Sub
    Sub DeleteXtraRowsCols(Wksht As Excel.Worksheet, EndTxt As String, LastUsedCol As Integer)
        'Delete column A, delete rows below last used row, delete colums right of LastUsedCol & delete hidden rows
        Dim MaxRowNo As Long, LastUsedRow As Long
        Dim MaxColNo As Long, RowNo As Long, TotRows As Long
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        MaxRowNo = Wksht.Rows.Count
        MaxColNo = Wksht.Columns.Count
        Wksht.Select()

        If Not Wksht.Cells.Find(EndTxt, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious) Is Nothing Then
            LastUsedRow = Wksht.Cells.Find(EndTxt, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious).Row
            Wksht.Range(XlSh.Cells(LastUsedRow + 1, 1), XlSh.Cells(MaxRowNo, 1)).EntireRow.Delete()
            Wksht.Range(XlSh.Cells(1, LastUsedCol + 1), XlSh.Cells(1, MaxColNo)).EntireColumn.Delete()
            Wksht.Columns("A:A").Delete
            TotRows = Wksht.UsedRange.Rows.Count
            For RowNo = 1 To TotRows
                If Wksht.Rows(RowNo).Hidden Then
                    Wksht.Rows(RowNo).Delete
                    TotRows = TotRows - 1
                    RowNo = RowNo - 1
                End If
            Next
            xlAp.ActiveWindow.FreezePanes = False
            xlAp.ActiveWindow.Split = False
            xlAp.ActiveWindow.ScrollRow = 1
            xlAp.ActiveWindow.SplitRow = 4
            xlAp.ActiveWindow.FreezePanes = True
        End If
    End Sub

End Module
