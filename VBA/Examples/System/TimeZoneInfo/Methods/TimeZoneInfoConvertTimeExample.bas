Attribute VB_Name = "TimeZoneInfoConvertTimeExample"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 23, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttime?view=netframework-4.8.1#system-timezoneinfo-converttime(system-datetime-system-timezoneinfo)

Option Explicit

Public Sub TimeZoneInfoConvertTime()
    Dim times() As IDateTime
    Objects.ToArray times, _
                    DateTime.CreateFromDateTime(2010, 1, 1, 0, 1, 0), _
                    DateTime.CreateFromDateTimeKind(2010, 1, 1, 0, 1, 0, DateTimeKind.DateTimeKind_Utc), _
                    DateTime.CreateFromDateTimeKind(2010, 1, 1, 0, 1, 0, DateTimeKind.DateTimeKind_Local), _
                    DateTime.CreateFromDateTime(2010, 11, 6, 23, 30, 0), _
                    DateTime.CreateFromDateTime(2010, 11, 7, 2, 30, 0)
                 
    TimeZoneInfo.ClearCachedData 'Clear incase timezone was changed.
    
    ' Retrieve the time zone for Eastern Standard Time (U.S. and Canada).
    Dim est As ITimeZoneInfo
    On Error Resume Next
    Set est = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")
    If Catch(TimeZoneNotFoundException) Then
        Debug.Print "Unable to retrieve the Eastern Standard time zone."
        Exit Sub
    ElseIf Catch(InvalidTimeZoneException) Then
        Debug.Print "Unable to retrieve the Eastern Standard time zone."
        Exit Sub
    End If
    On Error GoTo 0 'Stop code and display error
    
    ' Display the current time zone name.
    Debug.Print "Local time zone: " & TimeZoneInfo.Locale.DisplayName & VBA.vbNewLine
    ' Convert each time in the array.
    Dim varTimeToConvert As Variant
    For Each varTimeToConvert In times
        Dim timeToConvert As IDateTime
        Set timeToConvert = varTimeToConvert
        Dim targetTime As IDateTime
        Set targetTime = TimeZoneInfo.ConvertTime(timeToConvert, est)
        
        Debug.Print "Converted " & timeToConvert.ToString() & _
                     " " & DateTimeKindHelper.ToString(timeToConvert.Kind) & _
                     " to " & targetTime.ToString() & "."
    Next
End Sub

' The example displays the following output:
'    Local time zone: (GMT-08:00) Pacific Time (US & Canada)
'
'    Converted 1/1/2010 12:01:00 AM Unspecified to 1/1/2010 3:01:00 AM.
'    Converted 1/1/2010 12:01:00 AM Utc to 12/31/2009 7:01:00 PM.
'    Converted 1/1/2010 12:01:00 AM Local to 1/1/2010 3:01:00 AM.
'    Converted 11/6/2010 11:30:00 PM Unspecified to 11/7/2010 1:30:00 AM.
'    Converted 11/7/2010 2:30:00 AM Unspecified to 11/7/2010 5:30:00 AM.

