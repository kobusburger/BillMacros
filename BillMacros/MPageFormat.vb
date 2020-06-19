Module MPageFormat
    Structure HdrInfoType
        Dim PrevHdrRow As Integer 'Row of the previous, same level hdr
        Dim NoHdrItems As Integer 'No of non-empty items after and including the previous hdr
        Dim HNo As Integer 'No of the same level hdrs in the a hdr group
    End Structure
    Dim HdrInfo(3) As HdrInfoType '4 levels. The array index is the Hdr level starting at 0
    Dim NoShtItems As Integer    'Current No of non-empty items on the billsheet
    Dim HItNo As Integer 'Item no in the current hdr group

    Sub PageFormat()
        'On Error GoTo errHandler
        Dim Wksht As Excel.Worksheet, StartSht As Excel.Worksheet
        Dim FSSel As New FSheetSel
        Dim BillSheets As Excel.Sheets
        xlWb = xlAp.ActiveWorkbook
        xlSh = XlWb.ActiveSheet
        StartSht = XlSh
        ShowActivationNotice() 'Show activation warning window
        FSSel.Text = "Page Format"
        FSSel.ShowDialog()
        If FSSel.DialogResult <> System.Windows.Forms.DialogResult.OK Then Return

        CheckTemplateSheet("BillTemplate") 'Check BillTemplate sheet and named ranges and insert/ replace if not correct
        LogTrackInfo("PageFormat")
        BillSheets = XlWb.Worksheets
        If FSSel.SelSheets.Checked = True Then
            BillSheets = xlAp.ActiveWindow.SelectedSheets
        End If
        FSSel.Dispose()

        For Each Wksht In BillSheets
            Wksht.Select()

            If CheckSheetType(Wksht) = "#BillSheet#" Then
                EditFormatSub(Wksht)
                PageFormatSub(Wksht)
                If NoShtItems > 0 Then
                    Wksht.Tab.Color = Excel.XlRgbColor.rgbRed
                Else
                    Wksht.Tab.Color = Excel.XlRgbColor.rgbYellow
                End If
                If Not IsActivated() Then Exit For 'todo Changing Billsheets to a single sheet creates a runtime error
            End If
        Next
        StartSht.Select()
    End Sub
    Sub PageFormatSub(Billsheet As Excel.Worksheet)
        'This function does the following:
        '- Adds line spacing
        '- Adds page ends
        '- Sets print range
        '- Sets freeze panes
        Dim BillRow As Integer, EndBillRow As Integer
        Dim RowType As String
        XlWb = xlAp.ActiveWorkbook
        Dim BillTemplate As Excel.Worksheet = XlWb.Worksheets("BillTemplate")

        GetAllInfoPar() 'Put the page parameters on the Info sheet into BillInfoDict
        If CheckSheetType(Billsheet) = "#BillSheet#" Then
            xlAp.ScreenUpdating = False
            DeletePageBreaks(Billsheet)

            'Initialise HdrInfo & counter to zero
            SetHdrInfoToZero(0)
            NoShtItems = 0
            HItNo = 0

            With Billsheet
                .Rows.AutoFit()
                EndBillRow = .Columns(1).Find("#BillEnd#").Row
                BillRow = 1
                Do While BillRow <= EndBillRow 'Use a Do While because a For Next loop can be endless if the end value is changed
                    xlAp.StatusBar = "PageFormat/ Sheet: " & Billsheet.Name & "/ Row:" & BillRow & " of " & EndBillRow
                    If BillRow Mod 10 = 0 Then Windows.Forms.Application.DoEvents() 'DoEvents was added to avoid RuntimeCallableWrapper failed error
                    RowType = UCase(Trim(.Cells(BillRow, 1).value))
                    '                    DebugLog(Billsheet, BillRow)
                    Select Case RowType
                        Case "ITEM", "ITEM1", "ITEM2", "ITEM3" 'ITEM, ITEM1, ITEM2 or ITEM2 only has an effect on the formatting
                            If ItemIsNotEmpty(Billsheet, BillRow) Then
                                .Rows(BillRow).AutoFit
                                .Rows(BillRow + 1).Insert(shift:=Excel.XlDirection.xlDown)
                                BillTemplate.Range("Blank").Copy(.Cells(BillRow + 1, 1))
                                EndBillRow = EndBillRow + 1
                                IncrementNoItems()
                                CopyBillRow(Billsheet, BillTemplate.Range(RowType), .Cells(BillRow, 1))
                                BillRow = BillRow + 1
                            Else
                                .Rows(BillRow).Hidden = True
                            End If
                            BillRow = BillRow + BillTemplate.Range(RowType).Rows.Count

                        Case "IHDR", "IHDR1", "IHDR2", "IHDR3"
                            HideHdrGrpRows(Billsheet, BillRow, RowType)
                            If ItemIsNotEmpty(Billsheet, BillRow) Then
                                IncrementNoItems()
                            End If
                            CopyBillRow(Billsheet, BillTemplate.Range(RowType), .Cells(BillRow, 1))
                            .Rows(BillRow).AutoFit
                            .Rows(BillRow + 1).Insert(shift:=Excel.XlDirection.xlDown)
                            BillTemplate.Range("Blank").Copy(.Cells(BillRow + 1, 1))
                            EndBillRow = EndBillRow + 1
                            BillRow = BillRow + BillTemplate.Range(RowType).Rows.Count + 1 'todo Should multirow types be allowed?

                        Case "#BILLEND#"
                            HideHdrGrpRows(Billsheet, BillRow, RowType)
                            BillTemplate.Range("BILLEND").Copy(.Cells(BillRow, 1))
                            BillRow = BillRow + BillTemplate.Range("BILLEND").Rows.Count

                        Case "#BILLSHEET#"
                            'Only formats are copied
                            BillTemplate.Range("BILLSHEET").Copy()
                            .Cells(BillRow, 1).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteFormats, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
                            BillRow = BillRow + BillTemplate.Range("BILLSHEET").Rows.Count

                        Case "COLHDR"
                            BillTemplate.Range("COLHDR").Copy(.Cells(BillRow, 1))
                            BillRow = BillRow + BillTemplate.Range(RowType).Rows.Count

                        Case "NOTE"
                            CopyBillRow(Billsheet, BillTemplate.Range(RowType), .Cells(BillRow, 1))
                            .Rows(BillRow + 1).Insert(shift:=Excel.XlDirection.xlDown)
                            BillTemplate.Range("Blank").Copy(.Cells(BillRow + 1, 1))
                            EndBillRow = EndBillRow + 1
                            BillRow = BillRow + 1
                            BillRow = BillRow + BillTemplate.Range("NOTE").Rows.Count

                        Case "" 'Hide row for empty row type
                            .Rows(BillRow).Hidden = True
                            BillRow = BillRow + 1

                        Case Else 'Treat other row types as NOTE
                            .Rows(BillRow + 1).Insert(shift:=Excel.XlDirection.xlDown)
                            EndBillRow = EndBillRow + 1
                            BillRow = BillRow + 1
                            BillRow = BillRow + BillTemplate.Range("NOTE").Rows.Count
                    End Select
                Loop
                .PageSetup.PrintArea = .Range(.Cells(1, 2), .Cells(EndBillRow + 2, 8)).Address
                .Cells(1, 1).EntireColumn.Hidden = True
                .PageSetup.PrintArea = .Range(.Cells(1, 2), .Cells(EndBillRow + 2, 8)).Address
                If BillInfoDict.ContainsKey("PrintTitleRows") Then .PageSetup.PrintTitleRows = BillInfoDict("PrintTitleRows")
            End With
            SetForcedPagePar(Billsheet)   'Set the forced page parameters that affect page breaks
            InsertPageBreaks(Billsheet)
            FillLastPage(Billsheet)
        End If
        xlAp.ActiveWindow.FreezePanes = False
        xlAp.ActiveWindow.Split = False
        xlAp.ActiveWindow.ScrollRow = 1
        xlAp.ActiveWindow.ScrollColumn = 1
        xlAp.ActiveWindow.SplitColumn = 5
        xlAp.ActiveWindow.SplitRow = 4
        xlAp.ActiveWindow.FreezePanes = True

        xlAp.ScreenUpdating = True
        xlAp.StatusBar = False
    End Sub
    Function GetRowType(RowTypeText As String) As String
        RowTypeText = UCase(Trim(RowTypeText))
        GetRowType = RowTypeText
        Select Case RowTypeText
            Case "I0"
                GetRowType = "ITEM"
            Case "I1"
                GetRowType = "ITEM1"
            Case "I2"
                GetRowType = "ITEM2"
            Case "I3"
                GetRowType = "ITEM3"
            Case "H0"
                GetRowType = "IHDR"
            Case "H1"
                GetRowType = "IHDR1"
            Case "H2"
                GetRowType = "IHDR2"
            Case "H3"
                GetRowType = "IHDR3"
            Case "N"
                GetRowType = "NOTE"
        End Select
    End Function
    Sub InsertPageBreaks(Billsheet As Excel.Worksheet)
        Dim MaxPages As Integer 'The maximum number of pages that will be processed

        Dim BreakNo As Integer, TotalPageBreaks As Integer, BreakLine As Integer
        Dim PrintHeight As Single, PageHeight As Single
        Dim SumRowHeights As Single, PBRowsHeight As Single
        Dim BillTemplate As Excel.Worksheet = xlWb.Worksheets("BillTemplate")

        If IsActivated() Then
            MaxPages = 100 'Limit pages in case there is an endless loop
        Else
            MaxPages = 2 'Limit the number of page if Bill Macros is not activated
        End If

        'Determine the page breaks and insert pagebreak (PB) rows
        With Billsheet
            PBRowsHeight = BillTemplate.Range("PB").Height / 2
            PageHeight = 297 * 72 / 25.4    'A4 = 210 mm x 297 mm; convert to points
            PrintHeight = PageHeight - .PageSetup.TopMargin - .PageSetup.BottomMargin
            .ResetAllPageBreaks()

            'Add-page-breaks loop
            TotalPageBreaks = .HPageBreaks.Count
            BreakNo = 1
            While BreakNo <= TotalPageBreaks And BreakNo < MaxPages
                xlAp.StatusBar = "InsertPageBreaks/ Sheet: " & Billsheet.Name & "/ Break No:" & BreakNo & " of " & TotalPageBreaks
                If BreakNo Mod 10 = 0 Then Windows.Forms.Application.DoEvents() 'DoEvents was added to avoid RuntimeCallableWrapper failed error
                BreakLine = .HPageBreaks.Item(BreakNo).Location.Row 'Pagebreak is before pagebreak row
                SumRowHeights = 0
                'Decrement BreakLine until the sum of row heights is larger than the PB row height
                While SumRowHeights < PBRowsHeight
                    SumRowHeights += .Rows(BreakLine - 1).RowHeight
                    BreakLine -= 1
                End While

                'Insert 6 PB rows & manual page break
                .Rows(BreakLine - 1).Insert(shift:=Excel.XlDirection.xlDown)
                .Rows(BreakLine - 1).Insert(shift:=Excel.XlDirection.xlDown)
                .Rows(BreakLine - 1).Insert(shift:=Excel.XlDirection.xlDown)
                .Rows(BreakLine - 1).Insert(shift:=Excel.XlDirection.xlDown)
                .Rows(BreakLine - 1).Insert(Shift:=Excel.XlDirection.xlDown)
                .Rows(BreakLine - 1).Insert(shift:=Excel.XlDirection.xlDown)
                BillTemplate.Range("PB").Copy(.Cells(BreakLine - 1, 1))
                .Rows(BreakLine + 2).PageBreak = Excel.XlPageBreak.xlPageBreakManual
                BreakNo += 1
                TotalPageBreaks = .HPageBreaks.Count
            End While
            xlAp.StatusBar = False
        End With
    End Sub
    Sub FillLastPage(BillSheet As Excel.Worksheet)
        'Determine free space on the page and add empty blank row to fill the free space

        Dim NoOfExtraRows As Integer, PageFreeSpace As Single
        Dim LastBillRow As Integer, PrintTitleRowsHeight As Single
        Dim InsertRowHeight As Single, SumRowHeights As Single
        Dim PrevBreakRow As Integer, TotalPageBreaks As Integer
        Dim PrintHeight As Single
        Dim BillTemplate As Excel.Worksheet = xlWb.Worksheets("BillTemplate")

        With BillSheet
            TotalPageBreaks = .HPageBreaks.Count
            PrevBreakRow = 1
            PrintTitleRowsHeight = 0 'The first page does not have title rows
            If TotalPageBreaks > 0 Then 'If more than one page
                PrevBreakRow = .HPageBreaks.Item(TotalPageBreaks).Location.Row
                PrintTitleRowsHeight = .Range(.PageSetup.PrintTitleRows).Height
            End If

            Const ExtraSpaceToLeave = 1 'This is to allow for inaccurate printing heights
            LastBillRow = .Columns(1).Find("#BillEnd#").Row + 2
            PrintHeight = PagePrintHeight() - .PageSetup.TopMargin - .PageSetup.BottomMargin - PrintTitleRowsHeight - ExtraSpaceToLeave
            InsertRowHeight = 12.75 'For default font Calibri(Body Font) 10 (File/Options/General/When creating new workbooks)

            SumRowHeights = 0
            For i = PrevBreakRow To LastBillRow
                SumRowHeights += .Rows(i).RowHeight
            Next

            PageFreeSpace = PrintHeight - SumRowHeights
            NoOfExtraRows = Int(PageFreeSpace / InsertRowHeight)
            For i = 1 To NoOfExtraRows
                .Rows(LastBillRow - 2).Insert(shift:=Excel.XlDirection.xlDown) 'Insert uses the same font size and therefore default row height
                BillTemplate.Range("Blank").Copy(.Cells(LastBillRow - 2, 1)) 'Copy ensures the correct border lines
            Next
        End With
    End Sub
    Function PagePrintHeight() As Single
        'Printable page height in points

        'todo get printour height more accurate on last page
        'Test printer info System.Printing.PageImageableArea. https://stackoverflow.com/questions/296182/how-to-get-printer-info-in-net#296232
        'Sum of row heights is about 2.3% to 3.5% less than printout height depending on the border line thickness. Maybe increase PagePrintHeight with this %? 
        'Excel's Page layout ruler is not accurate. It changes with zoom. Print preview and printed pdf's seem accurate
        'https://docs.microsoft.com/en-us/office/troubleshoot/excel/worksheet-printed-different-size
        'https://inneka.com/programming/c/how-to-get-printer-info-in-net/

        PagePrintHeight = 297 * 72 / 25.4    'A4 = 210 mm x 297 mm; convert to points (1 point = 1/72 inch)
    End Function
    Sub DeletePageBreaks(Billsheet As Excel.Worksheet)
        With Billsheet
            Dim BillRow As Integer, LastBillRow As Integer
            LastBillRow = .Columns(1).Find("#BillEnd#").Row
            For BillRow = 1 To LastBillRow
                Select Case .Cells(BillRow, 1).value
                    Case "PB"
                        .Rows(BillRow).Delete
                        BillRow = BillRow - 1
                        LastBillRow = .Columns(1).Find("#BillEnd#").Row
                End Select
            Next
        End With
    End Sub
    Sub IncrementNoItems()
        'This should be called for each new measured item
        'The relevant counters will be incremented depending in which group the item is
        Dim HdLv As Integer, i As Integer, PrHdrRow As Integer
        NoShtItems = NoShtItems + 1
        HItNo = HItNo + 1
        PrHdrRow = xlAp.WorksheetFunction.Max(HdrInfo(0).PrevHdrRow, HdrInfo(1).PrevHdrRow, HdrInfo(2).PrevHdrRow, HdrInfo(3).PrevHdrRow)
        If PrHdrRow = 0 Then Exit Sub 'The item is not in an hdr group i.e. before any hdrs
        Select Case PrHdrRow
            Case HdrInfo(0).PrevHdrRow
                HdLv = 0
            Case HdrInfo(1).PrevHdrRow
                HdLv = 1
            Case HdrInfo(2).PrevHdrRow
                HdLv = 2
            Case HdrInfo(3).PrevHdrRow
                HdLv = 3
        End Select

        For i = 0 To HdLv 'Increment current and lower levels
            HdrInfo(i).NoHdrItems = HdrInfo(i).NoHdrItems + 1
        Next
    End Sub
    Sub HideHdrGrpRows(Billsheet As Excel.Worksheet, BillRow As Integer, RowType As String)
        'Hide rows if the hdr group is not used
        'Keep higher level hdr group if a lower hdr group is used
        'This sub should be called for each hdr and at Billend
        Dim i As Integer
        Select Case RowType
            Case "IHDR" 'H0 terminates groups H0, H1, H2 & H3 and starts new H0 group
                If HdrInfo(0).NoHdrItems = 0 And HdrInfo(0).PrevHdrRow > 0 Then 'Empty group and not first IHDR
                    HideRows(Billsheet, HdrInfo(0).PrevHdrRow, BillRow - 1)
                Else
                    HdrInfo(0).HNo = HdrInfo(0).HNo + 1
                End If
                If HdrInfo(1).NoHdrItems = 0 Then HideRows(Billsheet, HdrInfo(1).PrevHdrRow, BillRow - 1)
                If HdrInfo(2).NoHdrItems = 0 Then HideRows(Billsheet, HdrInfo(2).PrevHdrRow, BillRow - 1)
                If HdrInfo(3).NoHdrItems = 0 Then HideRows(Billsheet, HdrInfo(3).PrevHdrRow, BillRow - 1)
                SetHdrInfoToZero(1) 'Set higher levels to zero
                HdrInfo(0).PrevHdrRow = BillRow
                HdrInfo(0).NoHdrItems = 0 'Set item counter to zero
                HItNo = 0
            Case "IHDR1" 'H1 terminates groups H1, H2 & H3 and starts new H1 group
                If HdrInfo(1).NoHdrItems = 0 And HdrInfo(1).PrevHdrRow > 0 Then
                    HideRows(Billsheet, HdrInfo(1).PrevHdrRow, BillRow - 1)
                Else
                    HdrInfo(1).HNo = HdrInfo(1).HNo + 1
                End If
                If HdrInfo(2).NoHdrItems = 0 Then HideRows(Billsheet, HdrInfo(2).PrevHdrRow, BillRow - 1)
                If HdrInfo(3).NoHdrItems = 0 Then HideRows(Billsheet, HdrInfo(3).PrevHdrRow, BillRow - 1)
                SetHdrInfoToZero(2) 'Set higher levels to zero
                HdrInfo(1).NoHdrItems = 0 'Set item counter to zero
                HdrInfo(1).PrevHdrRow = BillRow
                HItNo = 0
            Case "IHDR2" 'H2 starts new H2 group and terminate H2 & H3 groups
                If HdrInfo(2).NoHdrItems = 0 And HdrInfo(2).PrevHdrRow > 0 Then
                    HideRows(Billsheet, HdrInfo(2).PrevHdrRow, BillRow - 1)
                Else
                    HdrInfo(2).HNo = HdrInfo(2).HNo + 1
                End If
                If HdrInfo(3).NoHdrItems = 0 Then HideRows(Billsheet, HdrInfo(3).PrevHdrRow, BillRow - 1)
                SetHdrInfoToZero(3) 'Set higher levels to zero
                HdrInfo(2).NoHdrItems = 0 'Set item counter to zero
                HdrInfo(2).PrevHdrRow = BillRow
                HItNo = 0
            Case "IHDR3" 'H3 starts new H3 group and terminate H3 group
                If HdrInfo(3).NoHdrItems = 0 And HdrInfo(3).PrevHdrRow > 0 Then
                    HideRows(Billsheet, HdrInfo(3).PrevHdrRow, BillRow - 1)
                Else
                    HdrInfo(3).HNo = HdrInfo(3).HNo + 1
                End If
                HdrInfo(3).PrevHdrRow = BillRow
                HdrInfo(3).NoHdrItems = 0 'Set item counter to zero
                HItNo = 0
            Case "#BILLEND#" 'BILLEND terminate all groups
                For i = 0 To 3 'Check all the levels
                    If HdrInfo(i).NoHdrItems = 0 And HdrInfo(i).PrevHdrRow > 0 Then
                        HideRows(Billsheet, HdrInfo(i).PrevHdrRow, BillRow - 1)
                    End If
                Next
        End Select
    End Sub
    Sub SetHdrInfoToZero(HdLv As Integer)
        'Set all values of HdLv and higher to zero
        Dim i As Integer
        For i = HdLv To 3
            HdrInfo(i).HNo = 0
            HdrInfo(i).NoHdrItems = 0
            HdrInfo(i).PrevHdrRow = 0
        Next
    End Sub
    Sub HideRows(Billsheet As Excel.Worksheet, ByVal FromRow As Integer, ByVal ToRow As Integer)
        'Hide all rows between StartRow and EndRow including StartRow and Endrow
        'Note that there are empty rows between each row and therefore the previous row is -2
        Dim R As Integer
        If ToRow > FromRow And FromRow > 1 Then
            For R = FromRow To ToRow
                Billsheet.Rows(R).Hidden = True
            Next
        End If
    End Sub
    Sub CopyBillRow(Billsheet As Excel.Worksheet, FromRange As Excel.Range, ToRange As Excel.Range)
        'Copy formats FromRange on BillTemplate to ToRange on Billsheet.
        'Formulas that start with "=" are modified and copied
        Dim ColOffset As Integer, RowOffset As Integer
        Dim LoopCell As Excel.Range

        xlAp.DisplayAlerts = False
        FromRange.Copy() 'Copy & paste formats
        ToRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteFormats, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
        For Each LoopCell In FromRange
            '            If xlAp.WorksheetFunction.IsFormula(LoopCell) Then
            If LoopCell.Text <> "" Then 'Only process cells that are not empty
                If LoopCell.Column = ItemNoCol Or LoopCell.Column = PricedAmtCol Then 'Only copy cells in the "Item No" and "Amount" columns
                    ColOffset = LoopCell.Column - FromRange.Column
                    RowOffset = LoopCell.Row - FromRange.Row
                    LoopCell.Copy()
                    ToRange.Offset(RowOffset, ColOffset).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteFormulas, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
                    ToRange.Offset(RowOffset, ColOffset).Formula = ReplaceCounters(ToRange.Offset(RowOffset, ColOffset).Formula)
                    If TypeOf ToRange.Offset(RowOffset, ColOffset).Value IsNot Int32 Then 'Only replace formulas in error free cells
                        ToRange.Offset(RowOffset, ColOffset).Formula = ReplaceFormulaRefs(ToRange.Offset(RowOffset, ColOffset).Formula, "'" & Billsheet.Name & "'!")
                    End If
                End If
            End If
        Next
        xlAp.DisplayAlerts = True
    End Sub
    Function ReplaceCounters(Formula As String) As String
        'Replace the hdr and item counters in the formula
        'todo This function also replace occurrances in strings. Is it possible to distinguish between variables and text?
        ReplaceCounters = UCase(Formula)
        ReplaceCounters = Replace(ReplaceCounters, "H0NO", HdrInfo(0).HNo, vbTextCompare)
        ReplaceCounters = Replace(ReplaceCounters, "H1NO", HdrInfo(1).HNo, vbTextCompare)
        ReplaceCounters = Replace(ReplaceCounters, "H2NO", HdrInfo(2).HNo, vbTextCompare)
        ReplaceCounters = Replace(ReplaceCounters, "H3NO", HdrInfo(3).HNo, vbTextCompare)
        ReplaceCounters = Replace(ReplaceCounters, "HITNO", HItNo, vbTextCompare)
    End Function
    Sub DebugLog(Wksh As Excel.Worksheet, Row As Integer)
        Dim HL As Integer
        Dim Col As Integer = 12
        Wksh.Cells(Row, Col) = "HINo=" & HItNo
        For HL = 0 To 3
            Wksh.Cells(Row, Col + 1 + 3 * HL) = HL & "HNo=" & HdrInfo(HL).HNo
            Wksh.Cells(Row, Col + 2 + 3 * HL) = HL & "INo=" & HdrInfo(HL).NoHdrItems
            Wksh.Cells(Row, Col + 3 + 3 * HL) = HL & "PrH=" & HdrInfo(HL).PrevHdrRow
        Next
    End Sub
End Module
