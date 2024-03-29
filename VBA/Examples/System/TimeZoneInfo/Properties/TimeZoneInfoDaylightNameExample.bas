Attribute VB_Name = "TimeZoneInfoDaylightNameExample"
'@Folder "Examples.System.TimeZoneInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.daylightname?view=netframework-4.8.1#examples

Option Explicit

''
' The following example defines a method named DisplayDateWithTimeZoneName that
' uses the IsDaylightSavingTime(DateTime) method to determine whether to display
' a time zone's standard time name or daylight saving time name.
''
Public Sub TimeZoneInfoDaylightName()
    Dim dateNow As DotNetLib.DateTime
    Set dateNow = DateTime.Now
    Dim localZone As DotNetLib.TimeZoneInfo
    Set localZone = TimeZoneInfo.Locale
    DisplayDateWithTimeZoneName dateNow, localZone
End Sub

Private Sub DisplayDateWithTimeZoneName(ByVal date1 As DotNetLib.DateTime, ByVal timeZone As DotNetLib.TimeZoneInfo)
    Debug.Print VBString.Format("The time is {0:t} on {0:d} {1}", _
                      date1, _
                      IIf(timeZone.IsDaylightSavingTime(date1), timeZone.DaylightName, timeZone.StandardName))
End Sub

' The example displays output similar to the following:
'    The time is 1:00 AM on 4/2/2006 Pacific Standard Time
