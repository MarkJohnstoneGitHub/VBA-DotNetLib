Attribute VB_Name = "TimeSpanTotalHoursExample"
'@Folder "Examples.System.TimeSpan.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.totalhours?view=netframework-4.8.1#examples

Option Explicit

''
' The following example instantiates a TimeSpan object and displays the value
' its TotalHours property. It also displays the value of each component
' (hours, minutes, seconds, and milliseconds) that forms the fractional part of
' the value of its TotalHours property.
''
Public Sub TimeSpanTotalHours()
    ' Define an interval of 1 day, 15+ hours.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.Create3(1, 15, 42, 45, 750)
    Debug.Print VBString.Format("Value of TimeSpan: {0}", interval)
    
    Debug.Print VBString.Format("{0:N5} hours, as follows:", interval.TotalHours)
    Debug.Print VBString.Format("   Hours:        {0,3}", _
                      interval.Days * 24 + interval.Hours)
    Debug.Print VBString.Format("   Minutes:      {0,3}", interval.Minutes)
    Debug.Print VBString.Format("   Seconds:      {0,3}", interval.Seconds)
    Debug.Print VBString.Format("   Milliseconds: {0,3}", interval.Milliseconds)
End Sub

' The example displays the following output:
'       Value of TimeSpan: 1.15:42:45.7500000
'       39.71271 hours, as follows:
'          Hours:         39
'          Minutes:       42
'          Seconds:       45
'          Milliseconds: 750

