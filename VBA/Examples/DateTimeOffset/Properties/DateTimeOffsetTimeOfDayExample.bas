Attribute VB_Name = "DateTimeOffsetTimeOfDayExample"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified July 19, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.timeofday?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the TimeOfDay property to extract the time and display it.")
Public Sub DateTimeOffsetTimeOfDay()
Attribute DateTimeOffsetTimeOfDay.VB_Description = "The following example uses the TimeOfDay property to extract the time and display it."
   Dim currentDate As DateTimeOffset
   Set currentDate = DateTimeOffset.CreateFromDateTimeParts(2008, 5, 10, 5, 32, 16, DateTimeOffset.Now.Offset)
   Dim currentTime As TimeSpan
   Set currentTime = currentDate.TimeOfDay
   Debug.Print "The current time is " & currentTime.ToString() & "."
   
' The example produces the following output:
'       The current time is 05:32:16.
End Sub
