Module MSetupPage
    Dim LastPageNo As Integer
    Dim BillInfoDict As New Dictionary(Of String, Object)
    Sub SetPage()
        'Set info parameters for each sheet
        'Set the First Page Number for each sheet
        Dim Wksht As Worksheet, BillSheets As Excel.Sheets 'Set freeze pane
        Dim ActShtName As String
        Dim xlAp As Excel.Application
        Dim XlWb As Excel.Workbook
        Dim XlSh As Excel.Worksheet
        Dim FSSel As New FSheetSel
        xlAp = Globals.ThisAddIn.Application
        XlWb = xlAp.ActiveWorkbook
        XlSh = XlWb.ActiveSheet
        ActShtName = XlSh.Name
        ShowActivationNotice() 'Show activation warning window

        FSSel.Text = "Setup page"
        FSSel.SelSheets.Enabled = True
        FSSel.Show()

        If FSSel.OK.Tag = True Then
            LogTrackInfo("SetPage")
            If FSSel.SelSheets.Enabled = True Then
                BillSheets = xlAp.ActiveWindow.SelectedSheets
            Else
                BillSheets = XlWb.Worksheets
            End If
            xlAp.ScreenUpdating = False
            GetInfoShtPar() 'Put the page parameters on the Info sheet into BillInfoDict
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

                    Call SetPageSub(Wksht)
                End If
            Next
        End If
        xlAp.ScreenUpdating = True
        xlAp.StatusBar = False
        XlWb.Sheets(ActShtName).Select
    End Sub
    Sub SetPageSub(Billsheet As Worksheet)
        'Set the parameters for "page layout" according to the info sheet
        'Add reference for "Microsoft Scrupting Library Runtime
        Dim Wksht As Worksheet
        Dim xlAp As Excel.Application
        xlAp = Globals.ThisAddIn.Application

        Call SetForcedPagePar(Billsheet) 'Set forced page parameters

        With Billsheet.PageSetup
            xlap.PrintCommunication = False

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
    Sub SetForcedPagePar(Billsheet As Worksheet)
        'The parameters are forced so that Bill Macros can work as expected
        Dim xlAp As Excel.Application
        xlAp = Globals.ThisAddIn.Application
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
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            xlAp.PrintCommunication = True
        End With
    End Sub
    Sub GetInfoShtPar()
        'Use Bill Info sheet and parameters if the sheet or parameters are avaialable
        'Otherwise insert Bill Info sheet
        Dim EndBillInfoRow As Integer, InfoRow As Integer
        Dim BillInfoSheet As Worksheet

        BillInfoSheet = GetInfoSheet()

        'Get parameters from Bill Info sheet
        EndBillInfoRow = BillInfoSheet.Columns(1).Find("#EndBillInfo#").Row
        For InfoRow = 2 To EndBillInfoRow
            If Len(BillInfoSheet.Cells(InfoRow, 1)) > 3 Then
                BillInfoDict(BillInfoSheet.Cells(InfoRow, 1).Value) = BillInfoSheet.Cells(InfoRow, 2).Value
            End If
        Next
    End Sub

End Module
