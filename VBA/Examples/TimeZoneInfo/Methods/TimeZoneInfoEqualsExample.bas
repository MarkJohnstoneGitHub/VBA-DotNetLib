Attribute VB_Name = "TimeZoneInfoEqualsExample"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 24, 2023
'@LastModified July 24, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.equals?view=netframework-4.8.1#system-timezoneinfo-equals(system-timezoneinfo)

Option Explicit

'@Description("The following example uses the Equals(TimeZoneInfo) method to determine whether the local time zone is Pacific Time or Eastern Time.")
Public Sub TimeZoneInfoEquals()
Attribute TimeZoneInfoEquals.VB_Description = "The following example uses the Equals(TimeZoneInfo) method to determine whether the local time zone is Pacific Time or Eastern Time."
    Dim thisTimeZone As TimeZoneInfo
    Dim zone1 As TimeZoneInfo
    Dim zone2 As TimeZoneInfo
    
    Set thisTimeZone = TimeZoneInfo.Locale
    Set zone1 = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time")
    Set zone2 = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")
    Debug.Print thisTimeZone.Equals(zone1)
    Debug.Print thisTimeZone.Equals(zone2)
End Sub
