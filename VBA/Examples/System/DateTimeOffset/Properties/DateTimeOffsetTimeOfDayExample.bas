Attribute VB_Name = "DateTimeOffsetTimeOfDayExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.timeofday?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the TimeOfDay property to extract the time and display it.
''
Public Sub DateTimeOffsetTimeOfDay()
    Dim currentDate As DotNetLib.DateTimeOffset
    Set currentDate = DateTimeOffset.CreateFromDateTimeParts(2008, 5, 10, 5, 32, 16, DateTimeOffset.Now.Offset)
    Dim currentTime As DotNetLib.TimeSpan
    Set currentTime = currentDate.TimeOfDay
    Debug.Print VBString.Format("The current time is {0}.", currentTime.ToString())
End Sub

' The example produces the following output:
'       The current time is 05:32:16.
