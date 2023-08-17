Attribute VB_Name = "TZInfoIsInvalidTimeExample"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 27, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.isinvalidtime?view=netframework-4.8.1#examples

Option Explicit

'In the Pacific Time zone, daylight saving time begins at 2:00 A.M. on April 2, 2006.
' The following code passes the time at one-minute intervals from 1:59 A.M. on
' April 2, 2006, to 3:01 A.M. on April 2, 2006, to the IsInvalidTime method of a
' TimeZoneInfo object that represents the Pacific Time zone.
' The console output indicates that all times from 2:00 A.M. on April 2, 2006, to
' 2:59 A.M. on April 2, 2006, are invalid.
Public Sub TimeZoneInfoIsInvalidTime()
    ' Specify DateTimeKind in Date constructor
    Dim baseTime As IDateTime
    Set baseTime = DateTime.CreateFromDateTimeKind(2007, 3, 11, 1, 59, 0, DateTimeKind.DateTimeKind_Unspecified)
    Dim newTime  As IDateTime
    
    ' Get Pacific Standard Time zone
    Dim pstZone As ITimeZoneInfo
    Set pstZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time")

    ' List possible invalid times for a 63-minute interval, from 1:59 AM to 3:01 AM
    Dim ctr As Long
    For ctr = 0 To 62
        ' Because of assignment, newTime.Kind is also DateTimeKind.Unspecified
        Set newTime = baseTime.AddMinutes(ctr)
        Debug.Print newTime.ToString() & " is invalid: " & _
        pstZone.IsInvalidTime(newTime)
    Next
End Sub

' Output:
'
'    11/03/2007 1:59:00 AM is invalid: False
'    11/03/2007 2:00:00 AM is invalid: True
'    11/03/2007 2:01:00 AM is invalid: True
'    11/03/2007 2:02:00 AM is invalid: True
'    11/03/2007 2:03:00 AM is invalid: True
'    11/03/2007 2:04:00 AM is invalid: True
'    11/03/2007 2:05:00 AM is invalid: True
'    11/03/2007 2:06:00 AM is invalid: True
'    11/03/2007 2:07:00 AM is invalid: True
'    11/03/2007 2:08:00 AM is invalid: True
'    11/03/2007 2:09:00 AM is invalid: True
'    11/03/2007 2:10:00 AM is invalid: True
'    11/03/2007 2:11:00 AM is invalid: True
'    11/03/2007 2:12:00 AM is invalid: True
'    11/03/2007 2:13:00 AM is invalid: True
'    11/03/2007 2:14:00 AM is invalid: True
'    11/03/2007 2:15:00 AM is invalid: True
'    11/03/2007 2:16:00 AM is invalid: True
'    11/03/2007 2:17:00 AM is invalid: True
'    11/03/2007 2:18:00 AM is invalid: True
'    11/03/2007 2:19:00 AM is invalid: True
'    11/03/2007 2:20:00 AM is invalid: True
'    11/03/2007 2:21:00 AM is invalid: True
'    11/03/2007 2:22:00 AM is invalid: True
'    11/03/2007 2:23:00 AM is invalid: True
'    11/03/2007 2:24:00 AM is invalid: True
'    11/03/2007 2:25:00 AM is invalid: True
'    11/03/2007 2:26:00 AM is invalid: True
'    11/03/2007 2:27:00 AM is invalid: True
'    11/03/2007 2:28:00 AM is invalid: True
'    11/03/2007 2:29:00 AM is invalid: True
'    11/03/2007 2:30:00 AM is invalid: True
'    11/03/2007 2:31:00 AM is invalid: True
'    11/03/2007 2:32:00 AM is invalid: True
'    11/03/2007 2:33:00 AM is invalid: True
'    11/03/2007 2:34:00 AM is invalid: True
'    11/03/2007 2:35:00 AM is invalid: True
'    11/03/2007 2:36:00 AM is invalid: True
'    11/03/2007 2:37:00 AM is invalid: True
'    11/03/2007 2:38:00 AM is invalid: True
'    11/03/2007 2:39:00 AM is invalid: True
'    11/03/2007 2:40:00 AM is invalid: True
'    11/03/2007 2:41:00 AM is invalid: True
'    11/03/2007 2:42:00 AM is invalid: True
'    11/03/2007 2:43:00 AM is invalid: True
'    11/03/2007 2:44:00 AM is invalid: True
'    11/03/2007 2:45:00 AM is invalid: True
'    11/03/2007 2:46:00 AM is invalid: True
'    11/03/2007 2:47:00 AM is invalid: True
'    11/03/2007 2:48:00 AM is invalid: True
'    11/03/2007 2:49:00 AM is invalid: True
'    11/03/2007 2:50:00 AM is invalid: True
'    11/03/2007 2:51:00 AM is invalid: True
'    11/03/2007 2:52:00 AM is invalid: True
'    11/03/2007 2:53:00 AM is invalid: True
'    11/03/2007 2:54:00 AM is invalid: True
'    11/03/2007 2:55:00 AM is invalid: True
'    11/03/2007 2:56:00 AM is invalid: True
'    11/03/2007 2:57:00 AM is invalid: True
'    11/03/2007 2:58:00 AM is invalid: True
'    11/03/2007 2:59:00 AM is invalid: True
'    11/03/2007 3:00:00 AM is invalid: False
'    11/03/2007 3:01:00 AM is invalid: False
'


