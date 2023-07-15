Attribute VB_Name = "TimeSpanTotalDaysExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.totaldays?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example instantiates a TimeSpan object and displays the value of its TotalDays property.")
'It also displays the value of each component (hours, minutes, seconds, milliseconds) that forms
'the fractional part of the value of its TotalDays property.
Public Sub TimeSpanTotalDays()
Attribute TimeSpanTotalDays.VB_Description = "The following example instantiates a TimeSpan object and displays the value of its TotalDays property."
   ' Define an interval of 3 days, 16+ hours.
   Dim interval As TimeSpan
   Set interval = TimeSpan.Create3(3, 16, 42, 45, 750)
   Debug.Print "Value of TimeSpan: " & interval.ToString
   'Console.WriteLine("Value of TimeSpan: {0}", interval)
   
   Debug.Print interval.TotalDays & " days, as follows:"
   Debug.Print "   Days:         " & interval.Days
   Debug.Print "   Hours:        " & interval.Hours
   Debug.Print "   Minutes:      " & interval.Minutes
   Debug.Print "   Seconds:      " & interval.Seconds
   Debug.Print "   Milliseconds: " & interval.Milliseconds
   
' The example displays the following output:
'       Value of TimeSpan: 3.16:42:45.7500000
'       3.69636 days, as follows:
'          Days:           3
'          Hours:         16
'          Minutes:       42
'          Seconds:       45
'          Milliseconds: 750
End Sub

