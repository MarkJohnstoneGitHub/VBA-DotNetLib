Attribute VB_Name = "TimeSpanTotalSecondsExample"
'@Folder "Examples.System.TimeSpan.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.totalseconds?view=netframework-4.8.1#examples

Option Explicit

''
' The following example instantiates a TimeSpan object and displays the value
' of its TotalSeconds property. It also displays the value of its milliseconds
' component, which forms the fractional part of the value of its TotalSeconds
' property.
''
Public Sub TimeSpanTotalSeconds()
    ' Define an interval of 1 day, 15+ hours.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.Create3(1, 15, 42, 45, 750)
    Debug.Print VBString.Format("Value of TimeSpan: {0}", interval)
    
    Debug.Print VBString.Format("{0:N5} seconds, as follows:", interval.totalSeconds)
    Debug.Print VBString.Format("   Seconds:      {0,8:N0}", _
                                interval.Days * 24 * 60 * 60 + _
                                interval.Hours * 60 * 60 + _
                                interval.Minutes * 60 + _
                                interval.Seconds)
    Debug.Print VBString.Format("   Milliseconds: {0,8}", interval.Milliseconds)
End Sub

' The example displays the following output:
'       Value of TimeSpan: 1.15:42:45.7500000
'       142,965.75000 seconds, as follows:
'          Seconds:       142,965
'          Milliseconds:      750
