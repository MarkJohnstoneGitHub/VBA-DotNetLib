Attribute VB_Name = "TimeZoneInfoDisplayNameExample"
'@Folder "Examples.System.TimeZoneInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.displayname?view=netframework-4.8.1#examples

Option Explicit

''
' The following example retrieves a TimeZoneInfo object that represents the
' local time zone and outputs its display name, standard time name, and
' daylight saving time name. The output is displayed for a system in the
' U.S. Pacific Standard Time zone.
' The output is displayed for a system in the U.S. Pacific Standard Time zone.
''
Public Sub TimeZoneInfoDisplayName()
    Dim localZone As DotNetLib.TimeZoneInfo
    Set localZone = TimeZoneInfo.Locale
    Debug.Print VBString.Format("Local Time Zone ID: {0}", localZone.Id)
    Debug.Print VBString.Format("   Display Name is: {0}.", localZone.DisplayName)
    Debug.Print VBString.Format("   Standard name is: {0}.", localZone.StandardName)
    Debug.Print VBString.Format("   Daylight saving name is: {0}.", localZone.DaylightName)
End Sub

' The example displays output like the following:
'     Local Time Zone ID: Pacific Standard Time
'        Display Name is: (UTC-08:00) Pacific Time (US & Canada).
'        Standard name is: Pacific Standard Time.
'        Daylight saving name is: Pacific Daylight Time.
