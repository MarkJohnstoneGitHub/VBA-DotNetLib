Attribute VB_Name = "TimeZoneInfoConvertTime3Example"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 23, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttime?view=netframework-4.8.1#system-timezoneinfo-converttime(system-datetime-system-timezoneinfo-system-timezoneinfo)

Option Explicit

''
' The following example illustrates the use of the
' ConvertTime(DateTime, TimeZoneInfo, TimeZoneInfo) method to convert from
' Hawaiian Standard Time to local time.
''
Public Sub TimeZoneInfoConvertTime3()
    Dim hwTime As DotNetLib.DateTime
    Set hwTime = DateTime.CreateFromDateTime(2007, 2, 1, 8, 0, 0)
    On Error Resume Next
    Dim hwZone As DotNetLib.TimeZoneInfo
    Set hwZone = TimeZoneInfo.FindSystemTimeZoneById("Hawaiian Standard Time")
    If Try() Then
        Debug.Print VBString.Format("{0} {1} is {2} local time.", _
                hwTime, _
                IIf(hwZone.IsDaylightSavingTime(hwTime), hwZone.DaylightName, hwZone.StandardName), _
                TimeZoneInfo.ConvertTime3(hwTime, hwZone, TimeZoneInfo.Locale))
    Else
        If Catch(TimeZoneNotFoundException) Then
            Debug.Print "The registry does not define the Hawaiian Standard Time zone."
        ElseIf Catch(InvalidTimeZoneException) Then
            Debug.Print "Registry data on the Hawaiian Standard Time zone has been corrupted."
        End If
    End If
    On Error GoTo 0 'Stop code and display error
End Sub
