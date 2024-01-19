Attribute VB_Name = "TZIFindSystemTimeZoneByIdEg"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 24, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.findsystemtimezonebyid?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the FindSystemTimeZoneById method to retrieve the
' Tokyo Standard Time zone. This TimeZoneInfo object is then used to convert
' the local time to the time in Tokyo and to determine whether it is Tokyo
' Standard Time or Tokyo Daylight Time.
''
Public Sub TimeZoneInfoFindSystemTimeZoneById()
    ' Get time in local time zone
    Dim thisTime As DotNetLib.DateTime
    Set thisTime = DateTime.Now
    Debug.Print VBString.Format("Time in {0} zone: {1}", _
                    IIf(TimeZoneInfo.Locale.IsDaylightSavingTime(thisTime), TimeZoneInfo.Locale.DaylightName, TimeZoneInfo.Locale.StandardName), _
                    thisTime)
    Debug.Print VBString.Format("   UTC Time: {0}", TimeZoneInfo.ConvertTimeToUtc2(thisTime, TimeZoneInfo.Locale))

    ' Get Tokyo Standard Time zone
    Dim tst As DotNetLib.TimeZoneInfo
    Set tst = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time")
    Dim tstTime As DotNetLib.DateTime
    Set tstTime = TimeZoneInfo.ConvertTime3(thisTime, TimeZoneInfo.Locale, tst)
    Debug.Print VBString.Format("Time in {0} zone: {1}", _
                    IIf(tst.IsDaylightSavingTime(thisTime), tst.DaylightName, tst.StandardName), _
                    tstTime)
    Debug.Print VBString.Format("   UTC Time: {0}", TimeZoneInfo.ConvertTimeToUtc2(tstTime, tst))
End Sub

' The example displays output like the following when run on a system in the
' U.S. Pacific Standard Time zone:
'       Time in Pacific Standard Time zone: 12/6/2013 10:57:51 AM
'          UTC Time: 12/6/2013 6:57:51 PM
'       Time in Tokyo Standard Time zone: 12/7/2013 3:57:51 AM
'          UTC Time: 12/6/2013 6:57:51 PM
