Attribute VB_Name = "TimeZoneInfoUtcExample"
'@Folder "Examples.System.TimeZoneInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.utc?view=netframework-4.8.1#examples

Option Explicit

''
' The following example retrieves a TimeZoneInfo object that represents
' Coordinated Universal Time (UTC) and outputs its display name, standard time
' name, and daylight saving time name.
''
Public Sub TimeZoneInfoUtc()
    Dim universalZone As DotNetLib.TimeZoneInfo
    Set universalZone = TimeZoneInfo.Utc
    Debug.Print VBString.Format("The universal time zone is {0}.", universalZone.DisplayName)
    Debug.Print VBString.Format("Its standard name is {0}.", universalZone.StandardName)
    Debug.Print VBString.Format("Its daylight savings name is {0}.", universalZone.DaylightName)
End Sub

' Output :
'
'    The universal time zone is UTC.
'    Its standard name is UTC.
'    Its daylight savings name is UTC.
