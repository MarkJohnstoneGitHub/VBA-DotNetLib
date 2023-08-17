Attribute VB_Name = "TimeZoneInfoUtcExample"
'@Folder "Examples.System.TimeZoneInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.utc?view=netframework-4.8.1#examples

Option Explicit

'The following example retrieves a TimeZoneInfo object that represents Coordinated Universal Time (UTC)
'and outputs its display name, standard time name, and daylight saving time name.
Public Sub TimeZoneInfoUtc()
    Dim universalZone As ITimeZoneInfo
    Set universalZone = TimeZoneInfo.Utc
    
    Debug.Print "The universal time zone is " & universalZone.DisplayName & "."
    Debug.Print "Its standard name is " & universalZone.StandardName & "."
    Debug.Print "Its daylight savings name is " & universalZone.DaylightName & "."
End Sub

' Output :
'
'    The universal time zone is UTC.
'    Its standard name is UTC.
'    Its daylight savings name is UTC.
