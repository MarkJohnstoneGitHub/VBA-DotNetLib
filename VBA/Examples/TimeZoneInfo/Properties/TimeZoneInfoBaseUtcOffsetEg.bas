Attribute VB_Name = "TimeZoneInfoBaseUtcOffsetEg"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.baseutcoffset?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the BaseUtcOffset property to display the difference between the local time and Coordinated Universal Time (UTC).")
Public Sub TimeZoneInfoBaseUtcOffset()
Attribute TimeZoneInfoBaseUtcOffset.VB_Description = "The following example uses the BaseUtcOffset property to display the difference between the local time and Coordinated Universal Time (UTC)."
    Dim localZone As ITimeZoneInfo
    Set localZone = TimeZoneInfo.Locale
    Debug.Print "The " & localZone.DisplayName & " time zone is " & _
                Abs(localZone.BaseUtcOffset.Hours) & ":" & _
                Abs(localZone.BaseUtcOffset.Minutes) & " " & _
                IIf((TimeSpan.GreaterThanOrEqual(localZone.BaseUtcOffset, TimeSpan.Zero)), "later", "earlier")
End Sub

'Output dependent on local settings:
' The (UTC+10:00) Canberra, Melbourne, Sydney time zone is 10:0 later
