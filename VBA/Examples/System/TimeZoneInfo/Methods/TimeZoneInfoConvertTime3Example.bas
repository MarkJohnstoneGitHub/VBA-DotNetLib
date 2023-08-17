Attribute VB_Name = "TimeZoneInfoConvertTime3Example"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 23, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttime?view=netframework-4.8.1#system-timezoneinfo-converttime(system-datetime-system-timezoneinfo-system-timezoneinfo)

Option Explicit

'@Description("The following example illustrates the use of the ConvertTime(DateTime, TimeZoneInfo, TimeZoneInfo) method to convert from Hawaiian Standard Time to local time.")
Public Sub TimeZoneInfoConvertTime3()
Attribute TimeZoneInfoConvertTime3.VB_Description = "The following example illustrates the use of the ConvertTime(DateTime, TimeZoneInfo, TimeZoneInfo) method to convert from Hawaiian Standard Time to local time."
    Dim hwTime As IDateTime
    Set hwTime = DateTime.CreateFromDateTime(2007, 2, 1, 8, 0, 0)
    On Error Resume Next
    Dim hwZone As ITimeZoneInfo
    Set hwZone = TimeZoneInfo.FindSystemTimeZoneById("Hawaiian Standard Time")
    If Try() Then
        Debug.Print hwTime.ToString() & " " & _
        IIf(hwZone.IsDaylightSavingTime(hwTime), hwZone.DaylightName, hwZone.StandardName) & " is " & _
        TimeZoneInfo.ConvertTime3(hwTime, hwZone, TimeZoneInfo.Locale).ToString() & " local time."
    Else
        If Catch(TimeZoneNotFoundException) Then
            Debug.Print "The registry does not define the Hawaiian Standard Time zone."
        ElseIf Catch(InvalidTimeZoneException) Then
            Debug.Print "Registry data on the Hawaiian Standard Time zone has been corrupted."
        End If
    End If
    On Error GoTo 0 'Stop code and display error
End Sub
