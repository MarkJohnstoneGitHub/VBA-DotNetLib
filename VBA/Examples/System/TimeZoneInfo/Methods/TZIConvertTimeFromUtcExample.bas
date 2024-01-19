Attribute VB_Name = "TZIConvertTimeFromUtcExample"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 23, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttimefromutc?view=netframework-4.8.1

Option Explicit

''
' The following example converts Coordinated Universal Time (UTC) to Central Time.
''
Public Sub TimeZoneInfoConvertTimeFromUtc()
    Dim timeUtc As DotNetLib.DateTime
    Set timeUtc = DateTime.UtcNow
    On Error Resume Next
    Dim cstZone As DotNetLib.TimeZoneInfo
    Set cstZone = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
    If Try() Then
        Dim cstTime As DotNetLib.DateTime
        Set cstTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, cstZone)
        Debug.Print VBString.Format("The date and time are {0} {1}.", _
                        cstTime, _
                        IIf(cstZone.IsDaylightSavingTime(cstTime), cstZone.DaylightName, cstZone.StandardName))
    ElseIf Catch(ArgumentException) Then
        Debug.Print "The registry does not define the Central Standard Time zone."
    ElseIf Catch(InvalidTimeZoneException) Then
        Debug.Print "Registry data on the Central Standard Time zone has been corrupted."
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

'Output:
'The date and time are 31/07/2023 6:23:03 AM Central Summer Time
