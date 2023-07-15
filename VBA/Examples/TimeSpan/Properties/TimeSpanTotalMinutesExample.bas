Attribute VB_Name = "TimeSpanTotalMinutesExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.totalminutes?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example instantiates a TimeSpan object and displays the value of its TotalMinutes property.")
' It also displays the value of each component (seconds, milliseconds) that forms the fractional
' part of the value of its TotalMinutes property.
Public Sub TimeSpanTotalMinutes()
Attribute TimeSpanTotalMinutes.VB_Description = "The following example instantiates a TimeSpan object and displays the value of its TotalMinutes property."
   ' Define an interval of 1 day, 15+ hours.
   Dim interval As TimeSpan
   Set interval = TimeSpan.Create3(1, 15, 42, 45, 750)
   Debug.Print "Value of TimeSpan: " & interval.ToString

   Debug.Print interval.TotalMinutes & " minutes, as follows:"
   Debug.Print "   Minutes:      " & interval.Days * 24 * 60 + _
                                     interval.Hours * 60 + _
                                     interval.Minutes
   Debug.Print "   Seconds:      " & interval.Seconds
   Debug.Print "   Milliseconds: " & interval.Milliseconds

'The example displays the following output:
'Value of TimeSpan: 1.15:42:45.7500000
'2382.7625 minutes, as follows:
'   Minutes:      2382
'   Seconds:      45
'   Milliseconds: 750
End Sub
