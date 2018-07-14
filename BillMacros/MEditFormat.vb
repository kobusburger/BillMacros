Module MEditFormat
    Sub EditFormat()
        Dim Wksht As Excel.Worksheet, BillSheets As Excel.Sheets
        Dim ActShtName As String
        Dim FSSel As New FSheetSel
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        ActShtName = XlSh.Name
        ShowActivationNotice() 'Show activation warning window
        FSSel.Text = "Edit Format"
        FSSel.ShowDialog()
        If FSSel.DialogResult <> System.Windows.Forms.DialogResult.OK Then Return

        LogTrackInfo("EditFormatVS")
        If FSSel.SelSheets.Checked = True Then
            BillSheets = xlAp.ActiveWindow.SelectedSheets
        Else
            BillSheets = XlWb.Worksheets
        End If
        FSSel.Dispose()

        For Each Wksht In BillSheets
            If CheckSheetType(Wksht) = "#BillSheet#" Then
                Wksht.Select()
                EditFormatSub(Wksht)
            End If
        Next
        XlWb.Sheets(ActShtName).Select
    End Sub
    Sub EditFormatSub(Billsheet As Excel.Worksheet)
        'This function does the following:
        '- Removes line spacing
        '- Removes page ends
        '- Deletes empty rows
        Dim BillRow As Integer, LastBillRow As Integer
        Dim xlAp As Excel.Application
        xlAp = Globals.ThisAddIn.Application
        If CheckSheetType(Billsheet) = "#BillSheet#" Then
            With Billsheet
                xlAp.ScreenUpdating = False
                LastBillRow = .Columns(1).Find("#BillEnd#").Row
                BillRow = 1
                Do While BillRow <= LastBillRow 'Use a Do While because a For Next loop can be endless if the end value is changed
                    xlAp.StatusBar = "EditFormat/ Sheet: " & Billsheet.Name & "/ Row:" & BillRow & " of " & LastBillRow
                    Select Case UCase(Trim(.Cells(BillRow, 1).value))
                        Case "PB"
                            .Rows(BillRow).Delete
                            BillRow = BillRow - 1
                            LastBillRow = LastBillRow - 1
                        Case Else
                            If xlAp.WorksheetFunction.CountA(.Rows(BillRow)) = 0 Then 'delete empty rows
                                .Rows(BillRow).Delete
                                BillRow = BillRow - 1
                                LastBillRow = LastBillRow - 1
                            End If
                    End Select
                    BillRow = BillRow + 1
                Loop
                .Cells(1, 1).EntireColumn.Hidden = False
                .Rows.AutoFit()
            End With
        End If
        xlAp.ScreenUpdating = True
        'freeze pane
        xlAp.ActiveWindow.FreezePanes = False
        xlAp.ActiveWindow.Split = False
        xlAp.ActiveWindow.ScrollRow = 1
        xlAp.ActiveWindow.ScrollColumn = 1
        xlAp.ActiveWindow.SplitColumn = 0
        xlAp.ActiveWindow.SplitRow = 4
        xlAp.ActiveWindow.FreezePanes = True

        xlAp.StatusBar = False
    End Sub


End Module
