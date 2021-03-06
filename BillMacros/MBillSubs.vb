﻿Module MBillSubs
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
        On Error GoTo 0
        NamedRangeExists = Err.Number = 0
    End Function
    Function ReplaceFormulaRefs(FormulaText As String, NewText As String) As String
        'Replaces all sheet references in FormulaText with NewText
        'Add VBA reference to "Microsoft VBScript Regular Expressions 5.5"
        'The following FormulaText can be used to test this function
        '='tt'!CAR&tt!E3&" ""X"" "&'[Bill macros.xla]BillTemplate'!$G$23 &" &'YY' "&'12 00'!G3&" "&'C:\Users\gert.brits\Documents\Bill of Quantities\[Bill macros.xla]BillTemplate'!$K$1
        'The syntax of a reference is "'path[workbookname]sheetname'!reference" but some references are not enclosed in single quotes

        'Add "If TypeOf Cell.Value IsNot Int32" before calling to avoid formula errors

        Dim regexpattern As String
        regexpattern = "['_a-zA-Z0-9\s\[\]\.:\\]+!"           '  "'(.*?)'" does not work because it only catches references in single quotes
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
        Dim BillSheets As Excel.Sheets
        Dim FormulasRange As Excel.Range
        xlWb = xlAp.ActiveWorkbook
        BillSheets = xlWb.Worksheets

        On Error Resume Next
        TemplateSheet = xlWb.Worksheets(TemplateName)
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
        TemplateSheet = xlWb.Worksheets(TemplateName)
        On Error Resume Next
        FormulasRange = TemplateSheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas)
        On Error GoTo 0
        'todo why is this done for sumtemplate & bill template if the sheet name does not change and info does not have formulas?
        If Not (FormulasRange Is Nothing) Then
            For Each Cell In FormulasRange
                If Cell.HasFormula Then
                    If TypeOf Cell.Value IsNot Int32 Then 'Only replace formulas in error free cells
                        Cell.Formula = ReplaceFormulaRefs(Cell.Formula, TemplateName & "!")
                    End If
                End If
            Next
        End If
    End Sub
    Sub CreateSheet(SheetName As String, ShtColor As Excel.XlRgbColor, AfterEnd As Boolean)
        'Delete SheetName and copy template sheet from BillMacrosTemplate
        Dim TemplateWB As Excel.Workbook
        Dim TemplatePath As String = IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.CodeBase)
        xlAp.ScreenUpdating = False 'Stop screen updating so that the second workbook does not show
        'It is not possible to copy worksheet objects between excel instances, only between workbooks in the same instance

        xlWb = xlAp.ActiveWorkbook
        xlAp.DisplayAlerts = False
        On Error Resume Next
        xlWb.Worksheets(SheetName).Delete()
        On Error GoTo 0
        xlAp.DisplayAlerts = True
        If Err.Number > 0 Then 'SheetName did not exist
            MsgBox(SheetName & " sheet was created", vbOKOnly)
        Else 'SheetName did exist
            MsgBox(SheetName & " sheet was recreated because it was not correct", vbOKOnly)
        End If
        TemplateWB = xlAp.Workbooks.Open(TemplatePath & "\" & BillMacrosTemplate, [ReadOnly]:=True)
        If AfterEnd = True Then 'Copy to last worksheet position
            TemplateWB.Worksheets(SheetName).Copy(After:=xlWb.Worksheets(xlWb.Worksheets.Count))
        Else 'Copy to first worksheet position
            TemplateWB.Worksheets(SheetName).Copy(before:=XlWb.Worksheets(1))
        End If
        XlWb.Worksheets(SheetName).Tab.Color = ShtColor
        TemplateWB.Close(False)
        xlAp.ScreenUpdating = True
    End Sub
    Sub DeleteBlankRows()
        'This function Deletes empty rows in the selected rows
        Dim RowNo As Integer, LastRow As Integer
        Dim StartRow As Integer
        Dim SelectedRows As Excel.Range
        xlWb = xlAp.ActiveWorkbook
        xlSh = xlWb.ActiveSheet
        SelectedRows = xlAp.Selection
        LastRow = SelectedRows.Rows.Count + SelectedRows.Row - 1
        If LastRow > xlSh.UsedRange.Rows.Count Then LastRow = xlSh.UsedRange.Rows.Count 'Limit LastRow to used range
        StartRow = SelectedRows.Row
        'xlAp.ScreenUpdating = False
        RowNo = StartRow
        While RowNo <= LastRow 'Use While because the variable of For may not be changed
            xlAp.StatusBar = "Row: " & RowNo & " of: " & LastRow
            If xlAp.WorksheetFunction.CountA(xlSh.Rows(RowNo)) = 0 Then 'delete empty rows
                xlSh.Rows(RowNo).entirerow.delete
                LastRow = LastRow - 1
            Else
                RowNo = RowNo + 1
            End If
            If RowNo Mod 10 = 0 Then Windows.Forms.Application.DoEvents() 'DoEvents was added to avoid RuntimeCallableWrapper failed error
        End While

        xlAp.ScreenUpdating = True
        xlAp.StatusBar = False
    End Sub

End Module
