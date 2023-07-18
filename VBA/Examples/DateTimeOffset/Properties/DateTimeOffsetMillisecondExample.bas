Attribute VB_Name = "DateTimeOffsetMillisecondExample"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified July 19, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.millisecond?view=netframework-4.8.1#examples

Option Explicit

'@Description("'The following example displays the number of milliseconds of a DateTimeOffset object by using a custom format specifier and by directly accessing the Millisecond property.")
Public Sub DateTimeOffsetMillisecond()
Attribute DateTimeOffsetMillisecond.VB_Description = "'The following example displays the number of milliseconds of a DateTimeOffset object by using a custom format specifier and by directly accessing the Millisecond property."
   Dim date1 As DateTimeOffset
   Set date1 = DateTimeOffset.CreateFromDateTimeParts2(2008, 3, 5, 5, 45, 35, 649, TimeSpan.Create(-7, 0, 0))
   Debug.Print "Milliseconds value of " & date1.ToString2("MM/dd/yyyy hh:mm:ss.fff") & " is " & date1.Millisecond & "."

' The example produces the following output:
'
' Milliseconds value of 03/05/2008 05:45:35.649 is 649.
End Sub
