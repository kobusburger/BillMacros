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

    Public Const MacroName As String = "Bill Macros.xla"
    Const VerYear As Integer = 2018,
    VerMonth As Integer = 7,
    VerDay As Integer = 16

    Const ActiveDays As Integer = 180 'The functionality will be reduced after the ActiveDays

    Sub AboutBill()
        Dim Msg As String
        Dim TerminationDate As Date
        Dim VersionDate As String
        VersionDate = VerYear & VerMonth.ToString("-00-") & VerDay.ToString("00")
        TerminationDate = DateSerial(VerYear, VerMonth, VerDay + ActiveDays)
        Msg = "The bill functions assist with the formatting of bills" & vbCrLf & vbCrLf &
    "Written by Kobus Burger © " & Left(VersionDate, 4) & vbCrLf &
    "083 228 9674 kobusgburger@gmail.com" & vbCrLf &
    vbCrLf & "Version date: " & VersionDate & vbCrLf &
    "Termination date: " & TerminationDate & vbCrLf &
    "Note that activity is being logged for statistical purposes"

        MsgBox(Msg, vbOKOnly, "Bill Macros")
    End Sub
    Sub LogTrackInfo(MenuItem As String)
        Dim TrackText As String
        Dim FileName As String
        Dim FilePath As String
        Dim UserName As String
        Dim DateTimeStr As String
        Dim VersionDate As String
        Dim LogTask As Threading.Tasks.Task
        '        Dim FSO As New FileSystemObject
        VersionDate = VerYear & VerMonth.ToString("-00-") & VerDay.ToString("00")
        'todo Maybe async (only VB, not VBA) can limit the delay if the file cannot be accessed
        'todo Implement logging on cloud server
        'https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/concepts/async/

        UserName = Environ$("Username")
        FilePath = "\\aurecon.info\shares\ZAPTA\Admin\Admin\GAUZABLD\2 Modify\Building Electrical Electronic Services\Software\ExcelAddins\"
        FileName = "tracking.txt"
        DateTimeStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        TrackText = DateTimeStr & vbTab &
            UserName & vbTab & vbTab &
            "BillMacrosVS" & vbTab & VersionDate & vbTab & MenuItem & vbCrLf
        '        On Error Resume Next
        '        TT = FSO.Drives.Item("N:") 'Drives correspond to Net Use. It may not show the correct connection status. This is specific to Aurecon network

        On Error Resume Next
        IO.File.AppendAllText(FilePath & FileName, TrackText)
        On Error GoTo 0
    End Sub
    Function IsActivated() As Boolean
        'Return true if Bill Macros is activated
        Dim TerminationDate As Date
        TerminationDate = DateSerial(VerYear, VerMonth, VerDay + ActiveDays)

        If DateDiff("d", Date.Now, TerminationDate) < 0 Then
            IsActivated = False
        Else
            IsActivated = True
        End If

    End Function
    Sub ShowActivationNotice()
        'Show warning windows
        Dim TerminationDate As Date, RemainingDays As Integer
        TerminationDate = DateSerial(VerYear, VerMonth, VerDay + ActiveDays)
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
