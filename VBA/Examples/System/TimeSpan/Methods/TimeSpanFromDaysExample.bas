Attribute VB_Name = "TimeSpanFromDaysExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.fromdays?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example creates several TimeSpan objects using the FromDays method.")
Public Sub TimeSpanFromDays()
Attribute TimeSpanFromDays.VB_Description = "The following example creates several TimeSpan objects using the FromDays method."
    Debug.Print VBAString.Unescape( _
            "This example of TimeSpan.FromDays( double )\n" + _
            "generates the following output.\n")
    Debug.Print VBAString.Format("{0,21}{1,18}", _
            "FromDays", "TimeSpan")
    Debug.Print VBAString.Format("{0,21}{1,18}", _
            "--------", "--------")

    Call GenTimeSpanFromDays(0.000000006)
    Call GenTimeSpanFromDays(0.000000017)
    Call GenTimeSpanFromDays(0.000123456)
    Call GenTimeSpanFromDays(1.234567898)
    Call GenTimeSpanFromDays(12345.678987654)
    Call GenTimeSpanFromDays(0.000011574)
    Call GenTimeSpanFromDays(0.000694444)
    Call GenTimeSpanFromDays(0.041666666)
    Call GenTimeSpanFromDays(1)
    Call GenTimeSpanFromDays(20.84745602)
End Sub

Private Sub GenTimeSpanFromDays(ByVal pDays As Double)
    ' Create a TimeSpan object and TimeSpan string from
    ' a number of days.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.FromDays(pDays)
    Dim timeInterval As String
    timeInterval = interval.ToString()
    Debug.Print VBAString.Format("{0,21}{1,26}", pDays, timeInterval)
End Sub

'/*
'This example of TimeSpan.FromDays( double )
'generates the following output.
'
'             FromDays          TimeSpan
'             --------          --------
'                6E-09          00:00:00.0010000
'              1.7E-08          00:00:00.0010000
'          0.000123456          00:00:10.6670000
'          1.234567898        1.05:37:46.6660000
'      12345.678987654    12345.16:17:44.5330000
'           1.1574E-05          00:00:01
'          0.000694444          00:01:00
'          0.041666666          01:00:00
'                    1        1.00:00:00
'          20.84745602       20.20:20:20.2000000
'*/
