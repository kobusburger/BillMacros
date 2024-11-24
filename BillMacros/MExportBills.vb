Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Module MExportBills
    Sub CreatePDF()
        'Save-as blank PDF
        'Only works in Excel 2007 and later

        Dim Wksht As Excel.Worksheet, StartSht As Excel.Worksheet
        Dim result As Boolean, First As Boolean
        xlWb = xlAp.ActiveWorkbook

        xlSh = xlWb.ActiveSheet
        StartSht = xlSh
        'ShowActivationNotice() 'Show activation warning window
        First = True
        For Each Wksht In xlWb.Worksheets
            Select Case Wksht.Tab.Color
                Case Excel.XlRgbColor.rgbRed 'Red = BillSheet
                    If First Then
                        Wksht.Select(True)
                        First = False
                    Else
                        Wksht.Select(False)
                    End If
                Case Excel.XlRgbColor.rgbGreen 'Green = Summary
                    Wksht.Select(False)
            End Select
            System.Windows.Forms.Application.DoEvents()
        Next
        result = xlAp.Dialogs.xlDialogSaveAs.Show(, 57) 'pdf type_num = 57
        StartSht.Select()
    End Sub
    Sub CreateStripped()
        'Delete hidden rows, delete non-bill columns & delete non-bill sheets
        Dim Wksht As Excel.Worksheet, FName As String
        xlWb = xlAp.ActiveWorkbook
        xlSh = xlWb.ActiveSheet
        CheckTemplateSheet("BillTemplate") 'Check BillTemplate sheet and named ranges and insert/ replace if not correct
        InitializeBillColNos()
        'Save bill with new name
        FName = Left(xlWb.Name, (InStrRev(xlWb.Name, ".", -1, vbTextCompare) - 1))
        If Not xlAp.Dialogs(Excel.XlBuiltInDialog.xlDialogSaveAs).Show(FName & " Stripped") Then Exit Sub
        For Each Wksht In xlWb.Worksheets
            xlAp.StatusBar = "Sheet: " & Wksht.Name
            xlAp.ScreenUpdating = False
            '            Wksht.Visible = Excel.XlSheetVisibility.xlSheetVisible 'Worksheets must be visible to avoid errors
            Select Case Wksht.Tab.Color
                Case Excel.XlRgbColor.rgbRed 'Red = BillSheet
                    ' Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas but this does not work with text that looks like dates
                    Wksht.UsedRange.Copy()
                    Wksht.UsedRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues,
                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
                    DeleteXtraRowsCols(Wksht, "#BillEnd#", AmtCol)
                Case Excel.XlRgbColor.rgbGreen 'Green = Summary
                    ' Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas but this does not work with text that looks like dates
                    Wksht.UsedRange.Copy()
                    Wksht.UsedRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues,
                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
                    DeleteXtraRowsCols(Wksht, "#SumEnd#", SumAmtCol)
                Case Else
                    xlAp.DisplayAlerts = False
                    Wksht.Delete()
                    xlAp.DisplayAlerts = True
            End Select
            xlAp.ScreenUpdating = True
            System.Windows.Forms.Application.DoEvents()
        Next
        xlAp.StatusBar = False
    End Sub
    Sub CreatePriced()
        'Delete hidden rows, delete non-bill columns, delete non-bill sheets & copy priced columns to bill
        Dim Wksht As Excel.Worksheet, FName As String
        Dim MaxRowNo As Integer, MaxColNo As Integer
        xlWb = xlAp.ActiveWorkbook
        xlSh = xlWb.ActiveSheet
        CheckTemplateSheet("BillTemplate") 'Check BillTemplate sheet and named ranges and insert/ replace if not correct
        InitializeBillColNos()
        'Save bill with new name
        FName = Left(xlWb.Name, (InStrRev(xlWb.Name, ".", -1, vbTextCompare) - 1))
        If Not xlAp.Dialogs(Excel.XlBuiltInDialog.xlDialogSaveAs).Show(FName & " Priced") Then Exit Sub
        xlAp.ScreenUpdating = False
        For Each Wksht In xlWb.Worksheets 'Do Summary (Green) first to preserve references to other sheets
            xlAp.StatusBar = "Sheet: " & Wksht.Name
            If Wksht.Tab.Color = Excel.XlRgbColor.rgbGreen Then
                ' Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas but this does not work with text that looks like dates
                Wksht.UsedRange.Copy()
                Wksht.UsedRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues,
                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
                MaxRowNo = Wksht.UsedRange.Rows.Count + 2
                MaxColNo = Wksht.UsedRange.Count + 2
                Wksht.Range(Wksht.Cells(1, SumPricedAmtCol), Wksht.Cells(MaxRowNo, SumPricedAmtCol)).Copy()
                Wksht.Cells(1, SumAmtCol).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=True, Transpose:=False)
                DeleteXtraRowsCols(Wksht, "#SumEnd#", SumAmtCol)
            End If
            System.Windows.Forms.Application.DoEvents()
        Next
        For Each Wksht In xlWb.Worksheets 'Do other sheets last
            xlAp.StatusBar = "Sheet: " & Wksht.Name
            Select Case Wksht.Tab.Color
                Case Excel.XlRgbColor.rgbRed 'Red = BillSheet
                    ' Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas but this does not work with text that looks like dates
                    Wksht.UsedRange.Copy()
                    Wksht.UsedRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues,
                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
                    MaxRowNo = Wksht.UsedRange.Rows.Count + 2
                    MaxColNo = Wksht.UsedRange.Count + 2
                    Wksht.Range(Wksht.Cells(1, PricedRateCol), Wksht.Cells(MaxRowNo, PricedAmtCol)).Copy()
                    Wksht.Cells(1, RateCol).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=True, Transpose:=False)
                    DeleteXtraRowsCols(Wksht, "#BillEnd#", AmtCol)
                Case Excel.XlRgbColor.rgbGreen 'Dummy for Green
                Case Else 'Delete unused sheets
                    xlAp.DisplayAlerts = False
                    Wksht.Delete()
                    xlAp.DisplayAlerts = True
            End Select
            System.Windows.Forms.Application.DoEvents()
        Next
        xlAp.ScreenUpdating = True
        xlAp.StatusBar = False
    End Sub
    Sub CreateTenderBill()
        'Delete hidden rows, delete non-bill columns, delete non-bill sheets, unprotect cells to be filled in & create formulas
        Dim Wksht As Excel.Worksheet, FName As String
        Dim MaxRowNo As Integer, MaxColNo As Integer, ColHDRRowNo As Integer
        Dim ContentRange As Excel.Range 'The range of the summary or bill sheet excluding the pricing columns
        Dim AmtRange As Excel.Range, RateRange As Excel.Range 'Rate and Amount column ranges
        Dim PricedAmtRange As Excel.Range

        xlWb = xlAp.ActiveWorkbook
        xlSh = xlWb.ActiveSheet
        CheckTemplateSheet("BillTemplate") 'Check BillTemplate sheet and named ranges and insert/ replace if not correct
        InitializeBillColNos()
        'Save bill with new name
        FName = Left(xlWb.Name, (InStrRev(xlWb.Name, ".", -1, vbTextCompare) - 1))
        If Not xlAp.Dialogs(Excel.XlBuiltInDialog.xlDialogSaveAs).Show(FName & " Tender") Then Exit Sub
        xlAp.ScreenUpdating = False
        For Each Wksht In xlWb.Worksheets 'Do Summary (Green) first to preserve references to other sheets
            xlAp.StatusBar = "Sheet: " & Wksht.Name
            If Wksht.Tab.Color = Excel.XlRgbColor.rgbGreen Then
                ' Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas but this does not work with text that looks like dates
                MaxRowNo = Wksht.UsedRange.Rows.Count + 2
                'MaxColNo = Wksht.UsedRange.Count + 2
                ContentRange = Wksht.Range(Wksht.Cells(1, 1), Wksht.Cells(MaxRowNo, SumAmtCol)) 'Summary range excluding price column

                'Remove formulas from SumRange and lock all cells
                ContentRange.Locked = True
                ContentRange.Copy()
                ContentRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues,
                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)

                'Copy price formulas to Amount column
                Wksht.Range(Wksht.Cells(1, SumPricedAmtCol), Wksht.Cells(MaxRowNo, SumPricedAmtCol)).Copy()
                AmtRange = Wksht.Range(Wksht.Cells(1, SumAmtCol), Wksht.Cells(MaxRowNo, SumAmtCol))
                AmtRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteFormulas,
                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)

                For Each cell In AmtRange 'Change billsheet references to defined name BillTotal
                    If InStr(1, cell.formula, "!") > 1 Then
                        cell.value = Left(cell.formula, InStr(1, cell.formula, "!")) & "BillTotal"
                    End If
                Next
                Wksht.Range(Wksht.Cells(ColHDRRowNo + 1, SumAmtCol), Wksht.Cells(MaxRowNo, SumAmtCol)).NumberFormat = "R#,##0.00"
                DeleteXtraRowsCols(Wksht, "#SumEnd#", SumAmtCol)
            End If
            System.Windows.Forms.Application.DoEvents()
        Next
        For Each Wksht In xlWb.Worksheets 'Do other sheets last
            xlAp.StatusBar = "Sheet: " & Wksht.Name

            Select Case Wksht.Tab.Color
                Case Excel.XlRgbColor.rgbRed 'Red = BillSheet
                    ' Wksht.UsedRange.Value = Wksht.UsedRange.Value 'Remove formulas but this does not work with text that looks like dates
                    MaxRowNo = Wksht.Columns("A").Find("#BillEnd#", SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious).row
                    ColHDRRowNo = Wksht.Columns("A").Find("ColHDR", SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious).row
                    'AmtRange = Wksht.Range(Wksht.Cells(ColHDRRowNo + 1, AmtCol), Wksht.Cells(MaxRowNo, AmtCol))
                    'RateRange = Wksht.Range(Wksht.Cells(ColHDRRowNo + 1, RateCol), Wksht.Cells(MaxRowNo, RateCol))
                    ContentRange = Wksht.Range(Wksht.Cells(1, 1), Wksht.Cells(MaxRowNo, AmtCol))
                    PricedAmtRange = Wksht.Range(Wksht.Cells(1, PricedAmtCol), Wksht.Cells(MaxRowNo, PricedAmtCol))

                    'Remove formulas and lock all cells
                    ContentRange.Locked = True
                    ContentRange.Copy()
                    ContentRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues,
                        Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
                    Wksht.Range(Wksht.Cells(ColHDRRowNo + 1, RateCol), Wksht.Cells(MaxRowNo, AmtCol)).NumberFormat = "R#,##0.00"
                    Wksht.Range(Wksht.Cells(ColHDRRowNo + 1, RateCol), Wksht.Cells(MaxRowNo, AmtCol)).HorizontalAlignment = XlHAlign.xlHAlignGeneral

                    'Insert amount formulas and unprotect rate cells
                    For row = ColHDRRowNo To MaxRowNo
                        If Wksht.Rows(row).hidden = False Then 'skip hidden rows
                            'Create page end/ start formulas
                            If (Wksht.Cells(row, 1).text = "PB") And Wksht.Cells(row, PricedAmtCol).hasformula Then
                                Wksht.Cells(row, AmtCol).formular1c1 = "=SUBTOTAL(9,R" & ColHDRRowNo + 1 & "C" & AmtCol & ":R" & row - 1 & "C" & AmtCol
                            End If
                            'Create bill end formula
                            If (Wksht.Cells(row, 1).text = "#BillEnd#") And Wksht.Cells(row, PricedAmtCol).hasformula Then
                                Wksht.Cells(row, AmtCol).formular1c1 = "=SUBTOTAL(9,R" & ColHDRRowNo + 1 & "C" & AmtCol & ":R" & row - 1 & "C" & AmtCol
                                Wksht.Names.Add("BillTotal", RefersTo:=Wksht.Cells(row, AmtCol))
                            End If
                            'Create item amount formulas
                            If (Wksht.Cells(row, UnitCol).text <> "") And (Wksht.Cells(row, AmtCol).text = "") And IsNumeric(Wksht.Cells(row, QtyCol).value) Then
                                Wksht.Cells(row, AmtCol).formular1c1 = "=R" & row & "C" & RateCol & "*R" & row & "C" & QtyCol
                                Wksht.Cells(row, RateCol).locked = False
                            End If
                            'Unlock non-numeric quantity items e.g. for rate only
                            If (Wksht.Cells(row, UnitCol).text <> "") And (Wksht.Cells(row, AmtCol).text <> "") And (IsNumeric(Wksht.Cells(row, QtyCol).value) = False) Then
                                Wksht.Cells(row, RateCol).locked = False
                            End If
                            'Set formats for percentage items
                            If Wksht.Cells(row, UnitCol).value = "%" Then
                                Wksht.Cells(row, QtyCol).NumberFormat = "R#,##0.00"
                                Wksht.Cells(row, RateCol).NumberFormat = "0.00%"
                            End If
                        End If 'not hidden row
                    Next
                    With ContentRange.FormatConditions.Add(Type:=XlFormatConditionType.xlExpression, Formula1:="=NOT(CELL(""protect"",A1))")
                        .interior.color = RGB(235, 241, 222) 'Light green
                    End With
                    DeleteXtraRowsCols(Wksht, "#BillEnd#", AmtCol)
                Case Excel.XlRgbColor.rgbGreen 'Dummy for Green
                Case Else 'Delete unused sheets
                    xlAp.DisplayAlerts = False
                    Wksht.Delete()
                    xlAp.DisplayAlerts = True
            End Select
            System.Windows.Forms.Application.DoEvents()
        Next
        xlAp.ScreenUpdating = True
        xlAp.StatusBar = False
    End Sub
    Sub DeleteXtraRowsCols(Wksht As Excel.Worksheet, EndTxt As String, LastUsedCol As Integer)
        'Delete column A, delete rows below last used row, delete colums right of LastUsedCol & delete hidden rows
        Dim MaxRowNo As Long, LastUsedRow As Long
        Dim MaxColNo As Long, RowNo As Long, TotRows As Long
        MaxRowNo = Wksht.UsedRange.Rows.Count + 2
        MaxColNo = Wksht.UsedRange.Columns.Count + 2
        Wksht.Select()

        If Not Wksht.Cells.Find(EndTxt, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious) Is Nothing Then
            LastUsedRow = Wksht.Cells.Find(EndTxt, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious).Row
            Wksht.Range(Wksht.Cells(LastUsedRow + 1, 1), Wksht.Cells(MaxRowNo, 1)).EntireRow.Delete()
            Wksht.Range(Wksht.Cells(1, LastUsedCol + 1), Wksht.Cells(1, MaxColNo)).EntireColumn.Delete()
            Wksht.Columns("A:A").Delete
            TotRows = Wksht.UsedRange.Rows.Count
            For RowNo = 1 To TotRows
                If Wksht.Rows(RowNo).Hidden Then
                    Wksht.Rows(RowNo).Delete
                    TotRows -= 1
                    RowNo -= 1
                End If
            Next
            xlAp.ActiveWindow.FreezePanes = False
            xlAp.ActiveWindow.Split = False
            'xlAp.ActiveWindow.Split = False
            'xlAp.ActiveWindow.ScrollRow = 1
            'xlAp.ActiveWindow.SplitRow = 4
            'xlAp.ActiveWindow.FreezePanes = True
        End If
    End Sub

End Module
