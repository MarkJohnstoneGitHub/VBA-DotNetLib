Attribute VB_Name = "TimeZoneInfoIsAmbiguousTimeEg"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 27, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.isambiguoustime?view=netframework-4.8.1

Option Explicit

''
' In the Pacific Time zone, daylight saving time ends at 2:00 A.M. on November 4, 2007.
' The following example passes the time at one-minute intervals from 12:59 A.M. on
' November 4, 2007, to 2:01 A.M. on November 4, 2007, to the IsAmbiguousTime(DateTime)
' method of a TimeZoneInfo object that represents the Pacific Time zone.
' The console output indicates that all times from 1:00 A.M. on November 4, 2007, to 1:59 A.M.
' on November 4, 2007, are ambiguous.
''
Public Sub TimeZoneInfoIsAmbiguousTime()
    ' Specify DateTimeKind in Date constructor
    Dim baseTime As DotNetLib.DateTime
    Set baseTime = DateTime.CreateFromDateTimeKind(2007, 11, 4, 0, 59, 0, DateTimeKind.DateTimeKind_Unspecified)
    Dim newTime  As DotNetLib.DateTime
    
    ' Get Pacific Standard Time zone
    Dim pstZone As DotNetLib.TimeZoneInfo
    Set pstZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time")

    ' List possible ambiguous times for 63-minute interval, from 12:59 AM to 2:01 AM
    Dim ctr As Long
    For ctr = 0 To 62
        ' Because of assignment, newTime.Kind is also DateTimeKind.Unspecified
        Set newTime = baseTime.AddMinutes(ctr)
        Debug.Print VBString.Format("{0} is ambiguous: {1}", newTime, pstZone.IsAmbiguousTime(newTime))
    Next
End Sub

' Output:
'    4/11/2007 12:59:00 AM is ambiguous: False
'    4/11/2007 1:00:00 AM is ambiguous: True
'    4/11/2007 1:01:00 AM is ambiguous: True
'    4/11/2007 1:02:00 AM is ambiguous: True
'    4/11/2007 1:03:00 AM is ambiguous: True
'    4/11/2007 1:04:00 AM is ambiguous: True
'    4/11/2007 1:05:00 AM is ambiguous: True
'    4/11/2007 1:06:00 AM is ambiguous: True
'    4/11/2007 1:07:00 AM is ambiguous: True
'    4/11/2007 1:08:00 AM is ambiguous: True
'    4/11/2007 1:09:00 AM is ambiguous: True
'    4/11/2007 1:10:00 AM is ambiguous: True
'    4/11/2007 1:11:00 AM is ambiguous: True
'    4/11/2007 1:12:00 AM is ambiguous: True
'    4/11/2007 1:13:00 AM is ambiguous: True
'    4/11/2007 1:14:00 AM is ambiguous: True
'    4/11/2007 1:15:00 AM is ambiguous: True
'    4/11/2007 1:16:00 AM is ambiguous: True
'    4/11/2007 1:17:00 AM is ambiguous: True
'    4/11/2007 1:18:00 AM is ambiguous: True
'    4/11/2007 1:19:00 AM is ambiguous: True
'    4/11/2007 1:20:00 AM is ambiguous: True
'    4/11/2007 1:21:00 AM is ambiguous: True
'    4/11/2007 1:22:00 AM is ambiguous: True
'    4/11/2007 1:23:00 AM is ambiguous: True
'    4/11/2007 1:24:00 AM is ambiguous: True
'    4/11/2007 1:25:00 AM is ambiguous: True
'    4/11/2007 1:26:00 AM is ambiguous: True
'    4/11/2007 1:27:00 AM is ambiguous: True
'    4/11/2007 1:28:00 AM is ambiguous: True
'    4/11/2007 1:29:00 AM is ambiguous: True
'    4/11/2007 1:30:00 AM is ambiguous: True
'    4/11/2007 1:31:00 AM is ambiguous: True
'    4/11/2007 1:32:00 AM is ambiguous: True
'    4/11/2007 1:33:00 AM is ambiguous: True
'    4/11/2007 1:34:00 AM is ambiguous: True
'    4/11/2007 1:35:00 AM is ambiguous: True
'    4/11/2007 1:36:00 AM is ambiguous: True
'    4/11/2007 1:37:00 AM is ambiguous: True
'    4/11/2007 1:38:00 AM is ambiguous: True
'    4/11/2007 1:39:00 AM is ambiguous: True
'    4/11/2007 1:40:00 AM is ambiguous: True
'    4/11/2007 1:41:00 AM is ambiguous: True
'    4/11/2007 1:42:00 AM is ambiguous: True
'    4/11/2007 1:43:00 AM is ambiguous: True
'    4/11/2007 1:44:00 AM is ambiguous: True
'    4/11/2007 1:45:00 AM is ambiguous: True
'    4/11/2007 1:46:00 AM is ambiguous: True
'    4/11/2007 1:47:00 AM is ambiguous: True
'    4/11/2007 1:48:00 AM is ambiguous: True
'    4/11/2007 1:49:00 AM is ambiguous: True
'    4/11/2007 1:50:00 AM is ambiguous: True
'    4/11/2007 1:51:00 AM is ambiguous: True
'    4/11/2007 1:52:00 AM is ambiguous: True
'    4/11/2007 1:53:00 AM is ambiguous: True
'    4/11/2007 1:54:00 AM is ambiguous: True
'    4/11/2007 1:55:00 AM is ambiguous: True
'    4/11/2007 1:56:00 AM is ambiguous: True
'    4/11/2007 1:57:00 AM is ambiguous: True
'    4/11/2007 1:58:00 AM is ambiguous: True
'    4/11/2007 1:59:00 AM is ambiguous: True
'    4/11/2007 2:00:00 AM is ambiguous: False
'    4/11/2007 2:01:00 AM is ambiguous: False


