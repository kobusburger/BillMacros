Module MBillSubs
    Dim xlAp As Excel.Application = Globals.ThisAddIn.Application
    Dim XlWb As Excel.Workbook
    Dim BillSheets As Excel.Sheets
    Dim XlSh As Excel.Worksheet
    Function ItemIsNotEmpty(Billsheet As Excel.Worksheet, ItemRow As Integer) As Boolean
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

    'Function GetInfoPar(InfoPar As String) As String
    '    Dim BillInfoSheet As Excel.Worksheet
    '    Dim EndBillInfoRow As Integer, InfoRow As Integer
    '    BillInfoSheet = GetInfoSheet()
    '    GetInfoPar = ""
    '    If BillInfoSheet Is Nothing Then Exit Function

    '    EndBillInfoRow = BillInfoSheet.Columns(1).Find("#EndBillInfo#").Row
    '    For InfoRow = 2 To EndBillInfoRow
    '        If BillInfoSheet.Cells(InfoRow, 1).value = InfoPar Then
    '            GetInfoPar = BillInfoSheet.Cells(InfoRow, 2).Value
    '        End If
    '    Next
    'End Function
    'Sub SetInfoPar(InfoPar As String, ParVal As Object)
    '    Dim BillInfoSheet As Excel.Worksheet
    '    Dim EndBillInfoRow As Integer, InfoRow As Integer
    '    BillInfoSheet = GetInfoSheet()
    '    If BillInfoSheet Is Nothing Then Exit Sub

    '    EndBillInfoRow = BillInfoSheet.Columns(1).Find("#EndBillInfo#").Row
    '    For InfoRow = 2 To EndBillInfoRow
    '        If BillInfoSheet.Cells(InfoRow, 1).value = InfoPar Then
    '            BillInfoSheet.Cells(InfoRow, 2).Value = ParVal
    '        End If
    '    Next
    'End Sub
    'Function GetInfoSheet() As Excel.Worksheet
    '    'Search for "Info" sheet and insert if it does not exist
    '    Dim Wksht As Excel.Worksheet
    '    Dim XlTemplate As Excel.Workbook

    '    GetInfoSheet = Nothing
    '    On Error Resume Next
    '    GetInfoSheet = XlWb.Worksheets("Info")
    '    On Error GoTo 0
    '    If GetInfoSheet Is Nothing Then
    '        xlAp.Workbooks(BillMacrosTemplate).Worksheets("Info").Copy(Before:=BillSheets(1))
    '        GetInfoSheet = XlWb.Worksheets("Info")
    '        MsgBox("The Info sheet was created because it did not exist", vbOKOnly)
    '    End If
    '    GetInfoSheet.Tab.Color = Excel.XlRgbColor.rgbBlue
    'End Function
    'Function GetBillTemplateSheet() As Excel.Worksheet
    '    'Search for "BillTemplate" sheet and insert if it does not exist
    '    Dim Wksht As Excel.Worksheet

    '    GetBillTemplateSheet = Nothing
    '    On Error Resume Next
    '    GetBillTemplateSheet = XlWb.Worksheets("BillTemplate")
    '    On Error GoTo 0
    '    If GetBillTemplateSheet Is Nothing Or Not CheckNamedRanges(BillSheets, "BillTemplate") Then
    '        xlAp.DisplayAlerts = False
    '        On Error Resume Next
    '        GetBillTemplateSheet.Delete()
    '        On Error GoTo 0
    '        xlAp.DisplayAlerts = True
    '        xlAp.Workbooks(BillMacrosTemplate).Worksheets("BillTemplate").Copy(Before:=BillSheets(1))
    '        GetBillTemplateSheet = XlWb.Worksheets("BillTemplate")
    '        MsgBox("The BillTemplate sheet was created because it did not exist or replaced because some of the named ranges do not exist", vbOKOnly)
    '    End If
    '    GetBillTemplateSheet.Tab.Color = Excel.XlRgbColor.rgbBlue
    'End Function
    Function CheckNamedRanges(BillSheets As Excel.Sheets, SheetName As String) As Boolean
        'Returns false if some range names in BillMacrosTemplate do not exist in the Bill Workbook template
        'It is assumed that SheetName exists
        Dim Rangename As Excel.Name
        Dim TemplateWB As Excel.Workbook
        Dim TemplatePath As String = IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.CodeBase)
        xlAp.ScreenUpdating = False 'Stop screen updating so that the second workbook does not show
        TemplateWB = xlAp.Workbooks.Open(TemplatePath & "\" & BillMacrosTemplate)
        'It is not possible to copy worksheet objects between excel instances, only between workbooks in the same instance

        CheckNamedRanges = True
        For Each Rangename In TemplateWB.Worksheets(SheetName).Names
            On Error Resume Next
            If BillSheets(SheetName).Names(Rangename.Name) Is Nothing Then
                CheckNamedRanges = False
                Exit Function
            End If
            On Error GoTo 0
        Next
        TemplateWB.Close(False)
        xlAp.ScreenUpdating = True
    End Function
    Function NamedRangeExists(Wksht As Excel.Worksheet, R As String) As Boolean
        'Returns true if the named range R exist on a Wksht
        'todo Is Named Ranges not also checked under CheckNamedRanges?
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

        'todo #REF creates a runtime error

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
    Sub CheckTemplateSheet(TemplateName As String)
        'Search for TemplateName sheet and insert if it does not exist
        'Update formula references to TemplateName
        Dim TemplateSheet As Excel.Worksheet
        Dim Cell As Excel.Range
        XlWb = xlAp.ActiveWorkbook
        BillSheets = XlWb.Worksheets

        On Error Resume Next
        TemplateSheet = XlWb.Worksheets(TemplateName)
        On Error GoTo 0
        If TemplateSheet Is Nothing Then
            CreateSheet(TemplateName, Excel.XlRgbColor.rgbBlue, False)
        ElseIf Not CheckNamedRanges(BillSheets, TemplateName) Then
            xlAp.DisplayAlerts = False
            On Error Resume Next
            TemplateSheet.Delete()
            On Error GoTo 0
            xlAp.DisplayAlerts = True
            CreateSheet(TemplateName, Excel.XlRgbColor.rgbBlue, False)
        End If

        'Replace sheet references in formules with TemplateName
        TemplateSheet = XlWb.Worksheets(TemplateName)
        On Error Resume Next
        For Each Cell In TemplateSheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas)
            If Cell.HasFormula Then
                Cell.Formula = ReplaceFormulaRefs(Cell.Formula, TemplateName & "!")
            End If
        Next
        On Error GoTo 0
    End Sub
    Sub CreateSheet(SheetName As String, ShtColor As Excel.XlRgbColor, AfterEnd As Boolean)
        'Delete SheetName and copy template sheet from BillMacrosTemplate
        Dim TemplateWB As Excel.Workbook
        Dim TemplatePath As String = IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.CodeBase)
        xlAp.ScreenUpdating = False 'Stop screen updating so that the second workbook does not show
        TemplateWB = xlAp.Workbooks.Open(TemplatePath & "\" & BillMacrosTemplate, [ReadOnly]:=True)
        'It is not possible to copy worksheet objects between excel instances, only between workbooks in the same instance

        On Error Resume Next
        xlAp.DisplayAlerts = False
        XlWb.Worksheets(SheetName).Delete()
        xlAp.DisplayAlerts = True
        If Err.Number > 0 Then 'SheetName did not exist
            MsgBox(SheetName & " sheet was created", vbOKOnly)
        Else 'SheetName did exist
            MsgBox(SheetName & " sheet was recreated because it was not correct", vbOKOnly)
        End If
        On Error GoTo 0
        If AfterEnd = True Then 'Copy to last worksheet position
            TemplateWB.Worksheets(SheetName).Copy(After:=XlWb.Worksheets(XlWb.Worksheets.Count))
        Else 'Copy to first worksheet position
            TemplateWB.Worksheets(SheetName).Copy(before:=XlWb.Worksheets(1))
        End If
        XlWb.Worksheets(SheetName).Tab.Color = ShtColor
        TemplateWB.Close(False)
        xlAp.ScreenUpdating = True
    End Sub
    Sub DeleteBlankRows()
        'This function Deletes empty rows in the selected rows
        Dim RangeRow As Integer, RowCount As Integer
        Dim RowNo As Integer, LastUsedRow As Integer
        Dim SelRows As Excel.Range
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        SelRows = xlAp.Selection
        LastUsedRow = xlAp.ActiveCell.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
        If LastUsedRow < SelRows(SelRows.Rows.Count).row Then
            RowCount = LastUsedRow - SelRows(1).row + 1 'limit row range to last used cell
        Else
            RowCount = SelRows.Rows.Count
        End If
        RangeRow = 1
        xlAp.ScreenUpdating = False
        While RangeRow <= RowCount
            RowNo = SelRows.Rows(RangeRow).row
            If xlAp.WorksheetFunction.CountA(XlSh.rows(RowNo)) = 0 Then 'delete empty rows
                XlSh.Rows(RowNo).entirerow.delete
                RowCount = RowCount - 1
            Else
                RangeRow = RangeRow + 1
            End If
        End While

        xlAp.ScreenUpdating = True
    End Sub

End Module
