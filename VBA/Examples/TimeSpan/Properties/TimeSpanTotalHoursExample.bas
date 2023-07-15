Attribute VB_Name = "TimeSpanTotalHoursExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.totalhours?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example instantiates a TimeSpan object and displays the value its TotalHours property.")
' It also displays the value of each component (hours, minutes, seconds, and milliseconds)
' that forms the fractional part of the value of its TotalHours property.
Public Sub TimeSpanTotalHours()
Attribute TimeSpanTotalHours.VB_Description = "The following example instantiates a TimeSpan object and displays the value its TotalHours property."
   ' Define an interval of 1 day, 15+ hours.
   Dim interval As TimeSpan
   Set interval = TimeSpan.Create3(1, 15, 42, 45, 750)
   Debug.Print "Value of TimeSpan: " & interval.ToString

   Debug.Print interval.TotalHours & " hours, as follows:"
   Debug.Print "   Hours:        " & interval.Days * 24 + interval.Hours
   Debug.Print "   Minutes:      " & interval.Minutes
   Debug.Print "   Seconds:      " & interval.Seconds
   Debug.Print "   Milliseconds: " & interval.Milliseconds

' The example displays the following output:
'       Value of TimeSpan: 1.15:42:45.7500000
'       39.71271 hours, as follows:
'          Hours:         39
'          Minutes:       42
'          Seconds:       45
'          Milliseconds: 750
End Sub
