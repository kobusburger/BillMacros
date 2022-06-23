﻿Module MEditFormat
    Sub EditFormat()
        Dim Wksht As Excel.Worksheet
        Dim ActShtName As String
        Dim FSSel As New FSheetSel
        Dim BillSheets As Excel.Sheets
        xlWb = xlAp.ActiveWorkbook
        xlSh = xlWb.ActiveSheet
        ActShtName = xlSh.Name
        'ShowActivationNotice() 'Show activation warning window
        FSSel.Text = "Edit Format"
        FSSel.ShowDialog()
        If FSSel.DialogResult <> System.Windows.Forms.DialogResult.OK Then Return

        LogTrackInfo("EditFormat")
        BillSheets = xlWb.Worksheets
        If FSSel.SelSheets.Checked = True Then
            BillSheets = xlAp.ActiveWindow.SelectedSheets
        End If
        FSSel.Dispose()

        xlAp.Application.Cursor = Excel.XlMousePointer.xlWait
        For Each Wksht In BillSheets
            If CheckSheetType(Wksht) = "#BillSheet#" Then
                Wksht.Select()
                EditFormatSub(Wksht)
            End If
        Next
        xlAp.Application.Cursor = Excel.XlMousePointer.xlDefault
        xlWb.Sheets(ActShtName).Select
    End Sub
    Sub EditFormatSub(Billsheet As Excel.Worksheet)
        'This function does the following:
        '- Removes line spacing
        '- Removes page ends
        '- Deletes empty rows
        Dim BillRow As Integer, LastBillRow As Integer, SplitRow As Long
        Dim BillTemplate As Excel.Worksheet = xlWb.Worksheets("BillTemplate")

        If CheckSheetType(Billsheet) = "#BillSheet#" Then
            With Billsheet
                xlAp.ScreenUpdating = False
                xlAp.Calculation = Excel.XlCalculation.xlCalculationManual
                LastBillRow = .Columns(1).Find("#BillEnd#").Row
                BillRow = 1
                Do While BillRow <= LastBillRow 'Use a Do While because a For Next loop can be endless if the end value is changed
                    xlAp.StatusBar = "EditFormat/ Sheet: " & Billsheet.Name & "/ Row:" & BillRow & " of " & LastBillRow
                    If BillRow Mod 10 = 0 Then Windows.Forms.Application.DoEvents() 'DoEvents was added to avoid RuntimeCallableWrapper failed error
                    Select Case UCase(Trim(.Cells(BillRow, 1).value))
                        Case "PB"
                            .Rows(BillRow).Delete
                            BillRow -= 1
                            LastBillRow -= 1
                        Case "COLHDR"
                            BillRow += BillTemplate.Range("COLHDR").Rows.Count
                            SplitRow = BillRow - 1
                        Case Else
                            If xlAp.WorksheetFunction.CountA(.Rows(BillRow)) = 0 Then 'delete empty rows
                                .Rows(BillRow).Delete
                                BillRow -= 1
                                LastBillRow -= 1
                            End If
                    End Select
                    BillRow += 1
                Loop
                .Cells(1, 1).EntireColumn.Hidden = False
                .Rows.AutoFit()
            End With
        End If
        xlAp.ScreenUpdating = True
        xlAp.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        'freeze pane
        xlAp.ActiveWindow.FreezePanes = False
        xlAp.ActiveWindow.Split = False
        xlAp.ActiveWindow.ScrollColumn = 1 'Left column number
        xlAp.ActiveWindow.SplitColumn = 0 'No of static columns
        xlAp.ActiveWindow.ScrollRow = 1 'Top row number
        xlAp.ActiveWindow.SplitRow = SplitRow 'No of static rows
        xlAp.ActiveWindow.FreezePanes = True

        xlAp.StatusBar = False
    End Sub


End Module
