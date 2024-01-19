Attribute VB_Name = "DateTimeOffsetUtcNowExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.utcnow?view=netframework-4.8.1#examples

Option Explicit

''
' The following example illustrates the relationship between Coordinated Universal
' Time (UTC) and local time.
''
Public Sub DateTimeOffsetUtcNow()
    Dim localTime As DotNetLib.DateTimeOffset
    Set localTime = DateTimeOffset.Now
    Dim utcTime As DotNetLib.DateTimeOffset
    Set utcTime = DateTimeOffset.UtcNow
    
    Debug.Print VBString.Format("Local Time:          {0}", localTime.ToString2("T"))
    Debug.Print VBString.Format("Difference from UTC: {0}", localTime.Offset.ToString())
    Debug.Print VBString.Format("UTC:                 {0}", utcTime.ToString2("T"))
End Sub

' If run on a particular date at 1:19 PM, the example produces
' the following output:
'    Local Time:          1:19:43 PM
'    Difference from UTC: -07:00:00
'    UTC:                 8:19:43 PM
