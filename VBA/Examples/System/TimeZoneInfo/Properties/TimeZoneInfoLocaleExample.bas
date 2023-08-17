Attribute VB_Name = "TimeZoneInfoLocaleExample"
'@Folder "Examples.System.TimeZoneInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.local?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example retrieves a TimeZoneInfo object that represents the local time zone and outputs its display name, standard time name, and daylight saving time name.")
Public Sub TimeZoneInfoLocale()
Attribute TimeZoneInfoLocale.VB_Description = "The following example retrieves a TimeZoneInfo object that represents the local time zone and outputs its display name, standard time name, and daylight saving time name."
    Dim localZone As ITimeZoneInfo
    Set localZone = TimeZoneInfo.Locale
    Debug.Print "Local Time Zone ID: " & localZone.Id
    Debug.Print "   Display Name is: " & localZone.DisplayName & "."
    Debug.Print "   Standard name is: " & localZone.StandardName & "."
    Debug.Print "   Daylight saving name is: " & localZone.DaylightName & "."
End Sub

' The example displays output like the following:
'     Local Time Zone ID: Pacific Standard Time
'        Display Name is: (UTC-08:00) Pacific Time (US & Canada).
'        Standard name is: Pacific Standard Time.
'        Daylight saving name is: Pacific Daylight Time.
