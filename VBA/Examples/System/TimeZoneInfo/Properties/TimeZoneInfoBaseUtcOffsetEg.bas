Attribute VB_Name = "TimeZoneInfoBaseUtcOffsetEg"
'@Folder "Examples.System.TimeZoneInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.baseutcoffset?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the BaseUtcOffset property to display the
' difference between the local time and Coordinated Universal Time (UTC).
''
Public Sub TimeZoneInfoBaseUtcOffset()
    Dim localZone As DotNetLib.TimeZoneInfo
    Set localZone = TimeZoneInfo.Locale
    Debug.Print VBString.Format("The {0} time zone is {1}:{2} {3} than Coordinated Universal Time.", _
                      localZone.DisplayName, _
                      Abs(localZone.BaseUtcOffset.Hours), _
                      Abs(localZone.BaseUtcOffset.Minutes), _
                      IIf((TimeSpan.GreaterThanOrEqual(localZone.BaseUtcOffset, TimeSpan.Zero)), "later", "earlier"))
End Sub

'Output dependent on local settings:
' The (UTC+10:00) Canberra, Melbourne, Sydney time zone is 10:0 later
