Attribute VB_Name = "TimeSpanTotalMillisecondsEg"
'@Folder "Examples.System.TimeSpan.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.totalmilliseconds?view=netframework-4.8.1#examples

Option Explicit

''
' The following example instantiates a TimeSpan object and displays the value
' of its TotalMilliseconds property.
''
Public Sub TimeSpanTotalMilliseconds()
    ' Define an interval of 1 day, 15+ hours.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.Create3(1, 15, 42, 45, 750)
    Debug.Print VBString.Format("Value of TimeSpan: {0}", interval)

    Debug.Print VBString.Format("There are {0:N5} milliseconds, as follows:", interval.TotalMilliseconds)
    Dim nMilliseconds As LongLong
    nMilliseconds = interval.Days * 24 * 60 * 60 * 1000 + _
                    interval.Hours * 60 * 60 * 1000 + _
                    interval.Minutes * 60 * 1000 + _
                    interval.Seconds * 1000 + _
                    interval.Milliseconds
    Debug.Print VBString.Format("   Milliseconds:     {0,18:N0}", nMilliseconds)
    Debug.Print VBString.Format("   Ticks:            {0,18:N0}", _
                      nMilliseconds * 10000 - interval.Ticks)
End Sub

' The example displays the following output:
'       Value of TimeSpan: 1.15:42:45.7500000
'       There are 142,965,750.00000 milliseconds, as follows:
'          Milliseconds:            142,965,750
'          Ticks:                             0

