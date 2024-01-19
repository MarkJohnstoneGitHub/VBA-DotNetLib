Attribute VB_Name = "TZISupportsDaylightSavingTimeEg"
'@Folder "Examples.System.TimeZoneInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.supportsdaylightsavingtime?view=netframework-4.8.1#examples

Option Explicit

''
' The following example retrieves a collection of all time zones that are
' available on a local system and displays the names of those that do not
' support daylight saving time.
''
Public Sub TimeZoneInfoSupportsDaylightSavingTime()
    Dim zones As DotNetLib.ReadOnlyCollection
    Set zones = TimeZoneInfo.GetSystemTimeZones()
    
    Dim varZone As Variant
    For Each varZone In zones
        Dim zone As DotNetLib.TimeZoneInfo
        Set zone = varZone
        If (Not zone.SupportsDaylightSavingTime) Then
            Debug.Print zone.DisplayName
        End If
    Next
End Sub

'Output:
'
'(UTC-12:00) International Date Line West
'(UTC-11:00) Coordinated Universal Time-11
'(UTC-10:00) Hawaii
'(UTC-09:30) Marquesas Islands
'(UTC-09:00) Coordinated Universal Time-09
'(UTC-08:00) Coordinated Universal Time-08
'(UTC-07:00) Arizona
'(UTC-06:00) Central America
'(UTC-06:00) Saskatchewan
'(UTC-05:00) Bogota, Lima, Quito, Rio Branco
'(UTC-04:00) Georgetown, La Paz, Manaus, San Juan
'(UTC-03:00) Cayenne, Fortaleza
'(UTC-02:00) Coordinated Universal Time-02
'(UTC-01:00) Cabo Verde Is.
'(UTC) Coordinated Universal Time
'(UTC+00:00) Monrovia, Reykjavik
'(UTC+01:00) West Central Africa
'(UTC+02:00) Harare, Pretoria
'(UTC+03:00) Kuwait, Riyadh
'(UTC+03:00) Nairobi
'(UTC+04:00) Abu Dhabi, Muscat
'(UTC+04:00) Tbilisi
'(UTC+04:30) Kabul
'(UTC+05:00) Ashgabat, Tashkent
'(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi
'(UTC+05:30) Sri Jayawardenepura
'(UTC+05:45) Kathmandu
'(UTC+06:00) Astana
'(UTC+06:30) Yangon (Rangoon)
'(UTC+07:00) Bangkok, Hanoi, Jakarta
'(UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi
'(UTC+08:00) Kuala Lumpur, Singapore
'(UTC+08:00) Taipei
'(UTC+08:45) Eucla
'(UTC+09:00) Osaka, Sapporo, Tokyo
'(UTC+09:00) Seoul
'(UTC+09:30) Darwin
'(UTC+10:00) Brisbane
'(UTC+10:00) Guam, Port Moresby
'(UTC+11:00) Solomon Is., New Caledonia
'(UTC+12:00) Coordinated Universal Time+12
'(UTC+13:00) Coordinated Universal Time+13
'(UTC+14:00) Kiritimati Island

