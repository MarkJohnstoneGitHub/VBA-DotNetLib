Attribute VB_Name = "DateTimeOffsetUtcNowExample"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified July 19, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.utcnow?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example illustrates the relationship between Coordinated Universal Time (UTC) and local time.")
Public Sub DateTimeOffsetUtcNow()
Attribute DateTimeOffsetUtcNow.VB_Description = "The following example illustrates the relationship between Coordinated Universal Time (UTC) and local time."
   Dim localTime As DateTimeOffset
   Set localTime = DateTimeOffset.Now
   Dim utcTime As DateTimeOffset
   Set utcTime = DateTimeOffset.UtcNow
   
   Debug.Print "Local Time:          " & localTime.ToString2("T")
   Debug.Print "Difference from UTC: " & localTime.Offset.ToString()
   Debug.Print "UTC:                 " & utcTime.ToString2("T")
   
' If run on a particular date at 1:19 PM, the example produces
' the following output:
'    Local Time:          1:19:43 PM
'    Difference from UTC: -07:00:00
'    UTC:                 8:19:43 PM
End Sub
