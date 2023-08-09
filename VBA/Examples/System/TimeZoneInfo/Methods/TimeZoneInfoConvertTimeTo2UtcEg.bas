Attribute VB_Name = "TimeZoneInfoConvertTimeTo2UtcEg"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 24, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttimetoutc?view=netframework-4.8.1#system-timezoneinfo-converttimetoutc(system-datetime-system-timezoneinfo)

Option Explicit

' The following example retrieves the current date from the local system and
' converts it to Coordinated Universal Time (UTC), then converts it to Tokyo
' Standard Time, and finally converts from Tokyo Standard Time back to UTC.
' Note that the two UTC times are identical.
Public Sub TimeZoneInfoConvertTimeToUtc2()
    Dim thisTime As IDateTime
    Set thisTime = DateTime.Now
    Debug.Print "Time in " & IIf(TimeZoneInfo.Locale.IsDaylightSavingTime(thisTime), TimeZoneInfo.Locale.DaylightName, TimeZoneInfo.Locale.StandardName) & _
                " zone: " & thisTime.ToString()
    Debug.Print "   UTC Time: " & TimeZoneInfo.ConvertTimeToUtc2(thisTime, TimeZoneInfo.Locale).ToString()
    
    ' Get Tokyo Standard Time zone
    Dim tst As ITimeZoneInfo
    Set tst = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time")
    Dim tstTime As IDateTime
    Set tstTime = TimeZoneInfo.ConvertTime3(thisTime, TimeZoneInfo.Locale, tst)
    Debug.Print "Time in " & IIf(tst.IsDaylightSavingTime(thisTime), tst.DaylightName, tst.StandardName) & _
                " zone: " & tstTime.ToString()
    Debug.Print "   UTC Time: " & TimeZoneInfo.ConvertTimeToUtc2(tstTime, tst).ToString()
End Sub

' The example displays output like the following when run on a system in the
' U.S. Pacific Standard Time zone:
'       Time in Pacific Standard Time zone: 12/6/2013 10:57:51 AM
'          UTC Time: 12/6/2013 6:57:51 PM
'       Time in Tokyo Standard Time zone: 12/7/2013 3:57:51 AM
'          UTC Time: 12/6/2013 6:57:51 PM
