Attribute VB_Name = "TimeSpanTotalDaysExample"
'@Folder "Examples.System.TimeSpan.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.totaldays?view=netframework-4.8.1#examples

Option Explicit

''
' The following example instantiates a TimeSpan object and displays the value
' of its TotalDays property. It also displays the value of each component
' (hours, minutes, seconds, milliseconds) that forms the fractional part of the
' value of its TotalDays property.
''
Public Sub TimeSpanTotalDays()
    ' Define an interval of 3 days, 16+ hours.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.Create3(3, 16, 42, 45, 750)
    Debug.Print VBString.Format("Value of TimeSpan: {0}", interval)
    
    Debug.Print VBString.Format("{0:N5} days, as follows:", interval.TotalDays)
    Debug.Print VBString.Format("   Days:         {0,3}", interval.Days)
    Debug.Print VBString.Format("   Hours:        {0,3}", interval.Hours)
    Debug.Print VBString.Format("   Minutes:      {0,3}", interval.Minutes)
    Debug.Print VBString.Format("   Seconds:      {0,3}", interval.Seconds)
    Debug.Print VBString.Format("   Milliseconds: {0,3}", interval.Milliseconds)
End Sub

' The example displays the following output:
'       Value of TimeSpan: 3.16:42:45.7500000
'       3.69636 days, as follows:
'          Days:           3
'          Hours:         16
'          Minutes:       42
'          Seconds:       45
'          Milliseconds: 750
