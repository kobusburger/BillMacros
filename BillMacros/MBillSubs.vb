Module MBillSubs
    Function ItemIsNotEmpty(Billsheet As Excel.Worksheet, ItemRow As Integer) As Boolean
        Dim xlAp As Excel.Application
        xlAp = Globals.ThisAddIn.Application
        ItemIsNotEmpty = False
        With Billsheet
            If xlAp.WorksheetFunction.CountA(.Range(.Cells(ItemRow, QtyCol), .Cells(ItemRow, AmtCol))) > 0 Then
                ItemIsNotEmpty = True
            End If
        End With
    End Function
    Function CheckSheetType(Billsheet As Excel.Worksheet) As String
        With Billsheet
            CheckSheetType = ""
            Select Case .Cells(1, 1).value
                Case "#BillSheet#"
                    If Not (.Columns(1).Find("#BillEnd#") Is Nothing) And
                        Not (.Columns(1).Find("ColHDR") Is Nothing) Then
                        CheckSheetType = "#BillSheet#"
                    Else
                        MsgBox("Column A of '" & Billsheet.Name & "' not correctly formatted")
                    End If
                Case "#SumSheet#"
                    If Not (.Columns(1).Find("BillGrpStart") Is Nothing) And
                        Not (.Columns(1).Find("BillGrpEnd") Is Nothing) Then
                        CheckSheetType = "#SumSheet#"
                    Else
                        MsgBox("Column A of '" & Billsheet.Name & "' not correctly formatted")
                    End If
                Case "#BillInfo#"
                    If Not (.Columns(1).Find("#EndBillInfo#") Is Nothing) Then
                        CheckSheetType = "#BillInfo#"
                    Else
                        MsgBox("Column A of '" & Billsheet.Name & "' not correctly formatted")
                    End If
                Case Else
            End Select
        End With
    End Function

    Function GetInfoPar(InfoPar As String) As String
        Dim BillInfoSheet As Excel.Worksheet
        Dim EndBillInfoRow As Integer, InfoRow As Integer
        BillInfoSheet = GetInfoSheet()
        GetInfoPar = ""
        If BillInfoSheet Is Nothing Then Exit Function

        EndBillInfoRow = BillInfoSheet.Columns(1).Find("#EndBillInfo#").Row
        For InfoRow = 2 To EndBillInfoRow
            If BillInfoSheet.Cells(InfoRow, 1).value = InfoPar Then
                GetInfoPar = BillInfoSheet.Cells(InfoRow, 2).Value
            End If
        Next
    End Function
    Sub SetInfoPar(InfoPar As String, ParVal As Object)
        Dim BillInfoSheet As Excel.Worksheet
        Dim EndBillInfoRow As Integer, InfoRow As Integer
        BillInfoSheet = GetInfoSheet()
        If BillInfoSheet Is Nothing Then Exit Sub

        EndBillInfoRow = BillInfoSheet.Columns(1).Find("#EndBillInfo#").Row
        For InfoRow = 2 To EndBillInfoRow
            If BillInfoSheet.Cells(InfoRow, 1).value = InfoPar Then
                BillInfoSheet.Cells(InfoRow, 2).Value = ParVal
            End If
        Next
    End Sub
    Function GetInfoSheet() As Excel.Worksheet
        'Search for "Info" sheet and insert if it does not exist
        Dim Wksht As Excel.Worksheet, BillSheets As Excel.Sheets
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        Dim XlTemplate As Excel.Workbook
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet

        BillSheets = XlWb.Worksheets
        GetInfoSheet = Nothing
        On Error Resume Next
        GetInfoSheet = XlWb.Worksheets("Info")
        On Error GoTo 0
        If GetInfoSheet Is Nothing Then 'todo NB!! not complete. Not working
            xlAp.Workbooks(BillMacrosTemplate).Worksheets("Info").Copy(Before:=BillSheets(1))
            GetInfoSheet = XlWb.Worksheets("Info")
            MsgBox("The Info sheet was created because it did not exist", vbOKOnly)
        End If
        GetInfoSheet.Tab.Color = Drawing.Color.Blue
    End Function
    Function GetBillTemplateSheet() As Excel.Worksheet
        'Search for "BillTemplate" sheet and insert if it does not exist
        Dim Wksht As Excel.Worksheet, BillSheets As Excel.Sheets
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet

        BillSheets = XlWb.Worksheets
        GetBillTemplateSheet = Nothing
        On Error Resume Next
        GetBillTemplateSheet = XlWb.Worksheets("BillTemplate")
        On Error GoTo 0
        If GetBillTemplateSheet Is Nothing Or Not CheckNamedRanges(BillSheets, "BillTemplate") Then
            xlAp.DisplayAlerts = False
            On Error Resume Next
            GetBillTemplateSheet.Delete()
            On Error GoTo 0
            xlAp.DisplayAlerts = True
            xlAp.Workbooks(BillMacrosTemplate).Worksheets("BillTemplate").Copy(Before:=BillSheets(1))
            GetBillTemplateSheet = XlWb.Worksheets("BillTemplate")
            MsgBox("The BillTemplate sheet was created because it did not exist or replaced because some of the named ranges do not exist", vbOKOnly)
        End If
        GetBillTemplateSheet.Tab.Color = Drawing.Color.Blue
    End Function
    Function GetSumTemplateSheet() As Excel.Worksheet
        'Search for "SumTemplate" sheet and insert if it does not exist
        Dim Wksht As Excel.Worksheet, BillSheets As Excel.Sheets
        Dim Cell As Excel.Range
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet

        BillSheets = XlWb.Worksheets
        GetSumTemplateSheet = Nothing
        On Error Resume Next
        GetSumTemplateSheet = XlWb.Worksheets("SumTemplate")
        On Error GoTo 0
        If GetSumTemplateSheet Is Nothing Or Not CheckNamedRanges(BillSheets, "SumTemplate") Then
            xlAp.DisplayAlerts = False
            On Error Resume Next
            GetSumTemplateSheet.Delete()
            On Error GoTo 0
            xlAp.DisplayAlerts = True
            xlAp.Workbooks(BillMacrosTemplate).Worksheets("SumTemplate").Copy(Before:=BillSheets(1))
            GetSumTemplateSheet = XlWb.Worksheets("SumTemplate")
            'Replace sheet references in formules with "SumTemplate"
            For Each Cell In GetSumTemplateSheet.Range("SumBillRow")
                If Cell.HasFormula Then
                    Cell.Formula = ReplaceFormulaRefs(Cell.Formula, "SumTemplate!")
                End If
            Next
            MsgBox("The SumTemplate sheet was created because it did not exist or replaced because some of the named ranges do not exist", vbOKOnly)
        End If
        GetSumTemplateSheet.Tab.Color = Drawing.Color.Blue
    End Function
    Function CheckNamedRanges(BillSheets As Excel.Sheets, SheetName As String) As Boolean 'Returns false if some range names do not exist
        Dim Rname As Excel.Name
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook

        CheckNamedRanges = True
        For Each Rname In XlWb.Worksheets(SheetName).Names
            On Error Resume Next 'Recreate the BillTemplate sheet if a named range does not exist
            If BillSheets(SheetName).Names(Rname.Name) Is Nothing Then
                CheckNamedRanges = False
                Exit Function
            End If
            On Error GoTo 0
        Next
    End Function
    Function GetSumSheet() As Excel.Worksheet
        'Search for "Summary" sheet and insert if it does not exist
        Dim Wksht As Excel.Worksheet, BillSheets As Excel.Sheets
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        BillSheets = XlWb.Worksheets
        GetSumSheet = Nothing
        On Error Resume Next
        GetSumSheet = BillSheets("Summary")
        On Error GoTo 0
        If GetSumSheet Is Nothing Then
            xlAp.Workbooks(BillMacrosTemplate).Worksheets("Summary").Copy(After:=BillSheets(BillSheets.Count))
            GetSumSheet = BillSheets(BillSheets.Count)
        End If
        GetSumSheet.Tab.Color = Drawing.Color.Green
    End Function
    Function NamedRangeExists(Wksht As Excel.Worksheet, R As String) As Boolean
        'Returns true if the named range R exist on a Wksht
        Dim TestR As Excel.Range
        On Error Resume Next
        TestR = Wksht.Range(R)
        NamedRangeExists = Err.Number = 0
        On Error GoTo 0
    End Function
    Function ReplaceFormulaRefs(FormulaText As String, NewText As String) As String
        'Replaces all sheet references in FormulaText with NewText
        'Add VBA reference to "Microsoft VBScript Regular Expressions 5.5"
        'The following FormulaText can be used to test this function
        '='tt'!CAR&tt!E3&" ""X"" "&'[Bill macros.xla]BillTemplate'!$G$23 &" &'YY' "&'12 00'!G3&" "&'C:\Users\gert.brits\Documents\Bill of Quantities\[Bill macros.xla]BillTemplate'!$K$1


        '------------------------------#REF creates a runtime error

        Dim regexpattern As String
        regexpattern = "['a-zA-Z0-9\s\[\]\.:\\]+!"
        Dim re As New System.Text.RegularExpressions.Regex(regexpattern)

        'Return all allowed charactors before "!" including "!"

        're.Global = True
        're.MultiLine = True

        ReplaceFormulaRefs = FormulaText
        '        If re.Test(FormulaText) Then
        ReplaceFormulaRefs = re.Replace(FormulaText, NewText)
        '        End If
    End Function


End Module
