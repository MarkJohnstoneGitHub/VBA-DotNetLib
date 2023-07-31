Attribute VB_Name = "TZIFindSystemTimeZoneByIdEg"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 24, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.findsystemtimezonebyid?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the FindSystemTimeZoneById method to retrieve the Tokyo Standard Time zone.")
'This TimeZoneInfo object is then used to convert the local time to the time in
'Tokyo and to determine whether it is Tokyo Standard Time or Tokyo Daylight Time.
Public Sub TimeZoneInfoFindSystemTimeZoneById()
Attribute TimeZoneInfoFindSystemTimeZoneById.VB_Description = "The following example uses the FindSystemTimeZoneById method to retrieve the Tokyo Standard Time zone."
    ' Get time in local time zone
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
                " zone: " & thisTime.ToString()
    Debug.Print "   UTC Time: " & TimeZoneInfo.ConvertTimeToUtc2(tstTime, tst).ToString()
End Sub

' The example displays output like the following when run on a system in the
' U.S. Pacific Standard Time zone:
'       Time in Pacific Standard Time zone: 12/6/2013 10:57:51 AM
'          UTC Time: 12/6/2013 6:57:51 PM
'       Time in Tokyo Standard Time zone: 12/7/2013 3:57:51 AM
'          UTC Time: 12/6/2013 6:57:51 PM
