Attribute VB_Name = "TimeZoneInfoStandardNameExample"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified July 22, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.standardname?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example defines a method named DisplayDateWithTimeZoneName that uses the IsDaylightSavingTime(DateTime) method to determine whether to display a time zone's standard time name or daylight saving time name.")
Public Sub TimeZoneInfoStandardName()
Attribute TimeZoneInfoStandardName.VB_Description = "The following example defines a method named DisplayDateWithTimeZoneName that uses the IsDaylightSavingTime(DateTime) method to determine whether to display a time zone's standard time name or daylight saving time name."
    Dim dateNow As DateTime
    Set dateNow = DateTime.Now
    Dim localZone As TimeZoneInfo
    Set localZone = TimeZoneInfo.Locale
    DisplayDateWithTimeZoneName dateNow, localZone
End Sub

Private Sub DisplayDateWithTimeZoneName(ByVal date1 As DateTime, ByVal timeZone As TimeZoneInfo)
    Debug.Print "The time is " & date1.ToString2("t") & " on " & _
                date1.ToString2("d") & " " & _
                IIf(timeZone.IsDaylightSavingTime(date1), timeZone.DaylightName, timeZone.StandardName)
End Sub

' The example displays output similar to the following:
'    The time is 1:00 AM on 4/2/2006 Pacific Standard Time