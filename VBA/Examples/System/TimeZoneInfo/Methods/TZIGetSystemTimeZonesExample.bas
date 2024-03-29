Attribute VB_Name = "TZIGetSystemTimeZonesExample"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 25, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.getsystemtimezones?view=netframework-4.8.1

Option Explicit

''
' The following example retrieves a collection of time zone objects that
' represent the time zones defined on a computer.
''
Public Sub TimeZoneInfoGetSystemTimeZones()
    Dim timeZones As ReadOnlyCollection
    Set timeZones = TimeZoneInfo.GetSystemTimeZones()
    
    Dim varTimeZone As Variant
    For Each varTimeZone In timeZones
        Dim timeZone As DotNetLib.TimeZoneInfo
        Set timeZone = varTimeZone
        Dim hasDST As Boolean
        hasDST = timeZone.SupportsDaylightSavingTime
        
        Debug.Print "ID: " & timeZone.Id
        Debug.Print "   Display Name: " & timeZone.DisplayName
        Debug.Print "   Daylight Name: " & timeZone.DaylightName
        Debug.Print IIf(hasDST, "   ***Has ", "   ***Does Not Have ") & "Daylight Saving Time***"
    Next
End Sub
