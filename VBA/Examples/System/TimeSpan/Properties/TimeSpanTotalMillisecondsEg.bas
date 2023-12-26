Attribute VB_Name = "TimeSpanTotalMillisecondsEg"
'@Folder "Examples.System.TimeSpan.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.totalmilliseconds?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example instantiates a TimeSpan object and displays the value of its TotalMilliseconds property.")
Public Sub TimeSpanTotalMilliseconds()
Attribute TimeSpanTotalMilliseconds.VB_Description = "The following example instantiates a TimeSpan object and displays the value of its TotalMilliseconds property."
   ' Define an interval of 1 day, 15+ hours.
   Dim interval As ITimeSpan
   Set interval = TimeSpan.Create3(1, 15, 42, 45, 750)
   Debug.Print "Value of TimeSpan: " & interval.ToString

   Debug.Print "There are " & interval.TotalMilliseconds & " milliseconds, as follows:"
   Dim nMilliseconds As LongLong
   nMilliseconds = interval.Days * 24 * 60 * 60 * 1000 + _
                   interval.Hours * 60 * 60 * 1000 + _
                   interval.Minutes * 60 * 1000 + _
                   interval.Seconds * 1000 + _
                   interval.Milliseconds
   
   Debug.Print "   Milliseconds:     " & nMilliseconds
   Debug.Print "   Ticks:            " & nMilliseconds * 10000 - interval.Ticks
End Sub

' The example displays the following output:
' Value of TimeSpan: 1.15:42:45.7500000
' There are 142965750 milliseconds, as follows:
'    Milliseconds:     142965750
'    Ticks:            0
