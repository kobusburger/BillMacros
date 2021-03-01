Module MGlobals
    Public Const _
    RefCol As Integer = 2,
    ItemNoCol As Integer = 3,
    DescrCol As Integer = 4,
    UnitCol As Integer = 5,
    QtyCol As Integer = 6,
    RateCol As Integer = 7,
    AmtCol As Integer = 8,
    PricedRateCol As Integer = 10,
    PricedAmtCol As Integer = 11

    Public Const _
    SumNoCol As Integer = 2,
    SumDesrCol As Integer = 3,
    SumAmtCol As Integer = 4,
    SumPricedAmtCol As Integer = 5

    Public Const BillMacrosTemplate As String = "BillMacrosTemplate.xlsx"
    'Const VerYear As Integer = 2019,
    'VerMonth As Integer = 7,
    'VerDay As Integer = 9

    Const ActiveDays As Integer = 360 'The functionality will be reduced after the ActiveDays

    'Global variable used in most modules
    Public xlAp As Excel.Application = Globals.ThisAddIn.Application
    Public xlWb As Excel.Workbook
    Public xlSh As Excel.Worksheet
    Public tc As New Microsoft.ApplicationInsights.TelemetryClient

    Sub AboutBill()
        Dim Msg As String
        Dim TerminationDate As Date
        Dim PublishVersion As String
        Dim AssemblyVersion As System.Version
        Dim VersionDate As New Date(2000, 1, 1)
        TerminationDate = VersionDate.AddDays(ActiveDays)
        If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            PublishVersion = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Else
            PublishVersion = ""
        End If
        AssemblyVersion = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version
        VersionDate = VersionDate.AddDays(AssemblyVersion.Build)
        VersionDate = VersionDate.AddSeconds(AssemblyVersion.Revision * 2)
        Msg = "The bill functions assist with the formatting of bills" & vbCrLf & vbCrLf &
            "Written by Kobus Burger © " & VersionDate.Year & vbCrLf &
            "083 228 9674 kobusgburger@gmail.com" & vbCrLf &
            vbCrLf & "Version date: " & VersionDate.ToString("yyyy-MM-dd HH:mm:ss") & vbCrLf &
            "Version: " & AssemblyVersion.ToString & vbCrLf &
            "Published version: " & PublishVersion '& GetDefaultPrinter()
        '"Note that activity is being logged for statistical purposes"
        '    "Termination date: " & TerminationDate & vbCrLf &

        MsgBox(Msg, vbOKOnly, "Bill Macros")
    End Sub
    Function GetDefaultPrinter() As String
        Dim settings As Drawing.Printing.PrinterSettings = New Drawing.Printing.PrinterSettings()
        Return settings.PrinterName
    End Function
    Sub LogTrackInfo(MenuItem As String) 'Use Azure application insights
        'https://carldesouza.com/how-to-create-custom-events-metrics-traces-in-azure-application-insights-using-c/
        'install the Microsoft.ApplicationInsights NuGet package
        Dim UserName As String
        Dim PubVer As String
        Dim EventProperties = New Dictionary(Of String, String)

        EventProperties.Add("FilePath", xlWb.FullName)
        UserName = Environ$("Username")
        PubVer = ""
        If Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            PubVer = Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4) 'Returns 4 components i.e. major.minor.build.revision
        End If

        tc.InstrumentationKey = "1ab23a13-3854-4c48-9bbb-5b1e2c7d9b2e"
        tc.Context.Session.Id = Guid.NewGuid.ToString
        tc.Context.Device.OperatingSystem = Environment.OSVersion.ToString
        tc.Context.User.AuthenticatedUserId = Environ$("Username")
        tc.Context.Component.Version = PubVer
        tc.TrackEvent(MenuItem, EventProperties)
        tc.Flush()
    End Sub
    Function IsActivated() As Boolean
        'Return true if Bill Macros is activated
        IsActivated = True 'Removed activation limit
        'Dim TerminationDate As Date
        'TerminationDate = DateSerial(VerYear, VerMonth, VerDay + ActiveDays)

        'If DateDiff("d", Date.Now, TerminationDate) < 0 Then
        '    IsActivated = False
        'Else
        '    IsActivated = True
        'End If

    End Function
    Sub ShowActivationNotice()
        'Show termination warning windows
        Dim TerminationDate As Date, RemainingDays As Integer
        Dim AssemblyVersion As System.Version
        Dim VersionDate As New Date(2000, 1, 1)
        AssemblyVersion = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version
        VersionDate = VersionDate.AddDays(AssemblyVersion.Build)
        VersionDate = VersionDate.AddSeconds(AssemblyVersion.Revision * 2)
        TerminationDate = VersionDate.AddDays(ActiveDays)
        RemainingDays = DateDiff("d", Date.Now, TerminationDate)

        Select Case RemainingDays
            Case Is < 0
                MsgBox("The termination date has passed and Bill Macros have reduced functionality" & vbCrLf &
            "The number of pages and sheets are limited in the Page Format function" & vbCrLf &
            "Please obtain an updated version.", vbOKOnly, "Bill Macros")
            Case 0 To 60
                MsgBox("The functionality of Bill Macros will be reduced in " & RemainingDays & " days." & vbCrLf &
            "The number of pages and sheets will be limited in the Page Format function" & vbCrLf &
            "Please obtain an updated version.", vbOKOnly, "Bill Macros")
        End Select

    End Sub
    Function OSInfo() As String 'Get Windows info
        'https://www.makeuseof.com/tag/see-pc-information-using-simple-excel-vba-script/
        'https://sites.google.com/site/beyondexcel/project-updates/exposingsystemsecretswithvbaandwmiapi
        'https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
        Dim OSys As Object
        Dim objos As Object
        On Error Resume Next
        ' Connect to WMI and obtain instances of Win32_OperatingSystem
        OSys = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
        For Each objos In OSys

            OSInfo = objos.SerialNumber
        Next

        If Err().Number <> 0 Then
            MsgBox(Err.Description)
            Err.Clear()
        End If
        On Error GoTo 0
    End Function
    'Sub TT() 'to test speed of checking if a network path exists
    '    Dim a As String, b As String, c As String
    '    Dim d, e, f
    '    Dim FSO As New FileSystemObject
    '    'Dir is sometimes fast but can sometimes be as slow as the other methods
    '    On Error Resume Next
    '    f = FSO.Drives
    '    d = FSO.Drives.Item("p:")
    '    e = Err.Number
    '    On Error GoTo 0
    'End Sub
End Module
