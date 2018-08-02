Module MSetupPage
    Dim LastPageNo As Integer
    Public BillInfoDict As New Dictionary(Of String, Object)
    Dim xlAp As Excel.Application = Globals.ThisAddIn.Application
    Dim XlWb As Excel.Workbook
    Dim XlSh As Excel.Worksheet
    Dim BillSheets As Excel.Sheets
    Sub SetPage()
        'Set info parameters for each sheet
        'Set the First Page Number for each sheet
        Dim Wksht As Excel.Worksheet 'Set freeze pane
        Dim ActShtName As String
        Dim FSSel As New FSheetSel
        XlWb = xlAp.ActiveWorkbook

        XlSh = XlWb.ActiveSheet
        ActShtName = XlSh.Name
        ShowActivationNotice() 'Show activation warning window

        FSSel.Text = "Setup page"
        FSSel.ShowDialog()
        If FSSel.DialogResult <> System.Windows.Forms.DialogResult.OK Then Return

        LogTrackInfo("SetPage")

        BillSheets = XlWb.Worksheets
        If FSSel.SelSheets.Checked = True Then
            BillSheets = xlAp.ActiveWindow.SelectedSheets
        End If
        FSSel.Dispose()

        xlAp.ScreenUpdating = False
        GetAllInfoPar() 'Put the page parameters on the Info sheet into BillInfoDict
        LastPageNo = 0
        If BillInfoDict.ContainsKey("FirstPageNumber") Then LastPageNo = BillInfoDict("FirstPageNumber") - 1
        For Each Wksht In BillSheets
            Wksht.Select()
            xlAp.StatusBar = "Setup Page/ Sheet: " & Wksht.Name
            If (CheckSheetType(Wksht) = "#BillSheet#" And Wksht.Tab.Color = RGB(255, 0, 0)) Or
            CheckSheetType(Wksht) = "#SumSheet#" Then
                Wksht.Select() 'The worksheet needs to be selected for pagesetup to accept changes correctly

                xlAp.PrintCommunication = False
                Wksht.PageSetup.FirstPageNumber = LastPageNo + 1
                LastPageNo = LastPageNo + Wksht.HPageBreaks.Count + 1
                xlAp.PrintCommunication = True

                SetPageSub(Wksht)
            End If
        Next
        xlAp.ScreenUpdating = True
        xlAp.StatusBar = False
        XlWb.Sheets(ActShtName).Select
    End Sub
    Sub SetPageSub(Billsheet As Excel.Worksheet)
        'Set the parameters for "page layout" according to the info sheet
        'Add reference for "Microsoft Scrupting Library Runtime

        SetForcedPagePar(Billsheet) 'Set forced page parameters

        With Billsheet.PageSetup
            xlAp.PrintCommunication = False

            If BillInfoDict.ContainsKey("PrintTitleRows") Then .PrintTitleRows = BillInfoDict("PrintTitleRows")

            'Margins tab
            If BillInfoDict.ContainsKey("LeftMargin") Then .LeftMargin = xlAp.CentimetersToPoints(BillInfoDict("LeftMargin"))
            If BillInfoDict.ContainsKey("RightMargin") Then .RightMargin = xlAp.CentimetersToPoints(BillInfoDict("RightMargin"))
            If BillInfoDict.ContainsKey("TopMargin") Then .TopMargin = xlAp.CentimetersToPoints(BillInfoDict("TopMargin"))
            If BillInfoDict.ContainsKey("BottomMargin") Then .BottomMargin = xlAp.CentimetersToPoints(BillInfoDict("BottomMargin"))
            If BillInfoDict.ContainsKey("HeaderMargin") Then .HeaderMargin = xlAp.CentimetersToPoints(BillInfoDict("HeaderMargin"))
            If BillInfoDict.ContainsKey("FooterMargin") Then .FooterMargin = xlAp.CentimetersToPoints(BillInfoDict("FooterMargin"))

            xlAp.PrintCommunication = True
            'PrintCommunication = False does not work for headers & footers
            'http://www.edugeek.net/blogs/pico/755-excel-2010-printer-fails-communicate.html
            If BillInfoDict.ContainsKey("LeftHeader") Then .LeftHeader = BillInfoDict("LeftHeader")
            If BillInfoDict.ContainsKey("CenterHeader") Then .CenterHeader = BillInfoDict("CenterHeader")
            If BillInfoDict.ContainsKey("RightHeader") Then .RightHeader = BillInfoDict("RightHeader")
            If BillInfoDict.ContainsKey("LeftFooter") Then .LeftFooter = BillInfoDict("LeftFooter")
            If BillInfoDict.ContainsKey("CenterFooter") Then .CenterFooter = BillInfoDict("CenterFooter")
            If BillInfoDict.ContainsKey("RightFooter") Then .RightFooter = BillInfoDict("RightFooter")

        End With
    End Sub
    Sub SetForcedPagePar(Billsheet As Excel.Worksheet)
        'The parameters are forced so that Bill Macros can work as expected
        With Billsheet.PageSetup
            xlAp.PrintCommunication = False
            .PrintTitleColumns = ""
            .Zoom = 100
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlPortrait
            .Draft = False
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .PrintHeadings = False
            .PrintGridlines = False
            '.PrintComments = Excel.XlPrintLocation.xlPrintNoComments 'This gives an error when PrintCommunications is set to Yes
            xlAp.PrintCommunication = True 'An error here refers to a problem in the previous lines
        End With
    End Sub
    Sub GetAllInfoPar()
        'Insert all Info sheet parameters in BillInfoDict
        'Create Info sheet if it does not exists
        Dim EndBillInfoRow As Integer, InfoRow As Integer
        Dim BillInfoSheet As Excel.Worksheet
        XlWb = xlAp.ActiveWorkbook

        CheckTemplateSheet("Info") 'Check Info sheet and named ranges and insert/ replace if not correct
        BillInfoSheet = XlWb.Worksheets("Info")

        'Get parameters from Bill Info sheet
        EndBillInfoRow = BillInfoSheet.Columns(1).Find("#EndBillInfo#").Row
        For InfoRow = 2 To EndBillInfoRow
            If Len(BillInfoSheet.Cells(InfoRow, 1).value) > 3 Then 'todo .Value is not recognised by VS because of late binding. VS does not know what type Cells(1,1) is.
                BillInfoDict(BillInfoSheet.Cells(InfoRow, 1).Value) = BillInfoSheet.Cells(InfoRow, 2).Value
            End If
        Next
    End Sub

End Module
