Attribute VB_Name = "TimeSpanTotalSecondsExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.totalseconds?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example instantiates a TimeSpan object and displays the value of its TotalSeconds property.")
' It also displays the value of its milliseconds component, which forms the fractional part
' of the value of its TotalSeconds property.
Public Sub TimeSpanTotalSeconds()
Attribute TimeSpanTotalSeconds.VB_Description = "The following example instantiates a TimeSpan object and displays the value of its TotalSeconds property."
   ' Define an interval of 1 day, 15+ hours.
   Dim interval As TimeSpan
   Set interval = TimeSpan.Create3(1, 15, 42, 45, 750)
   Debug.Print "Value of TimeSpan: " & interval.ToString

   Debug.Print interval.TotalSeconds & " seconds, as follows:"
   Debug.Print "   Seconds:      " & interval.Days * 24 * 60 * 60 + _
                                     interval.Hours * 60 * 60 + _
                                     interval.Minutes * 60 + _
                                     interval.Seconds
   Debug.Print "   Milliseconds: " & interval.Milliseconds

' The example displays the following output:
' Value of TimeSpan: 1.15:42:45.7500000
' 142965.75 seconds, as follows:
'    Seconds:      142965
'    Milliseconds: 750
End Sub
