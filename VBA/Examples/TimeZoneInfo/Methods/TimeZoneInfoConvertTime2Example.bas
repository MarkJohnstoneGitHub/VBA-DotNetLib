Attribute VB_Name = "TimeZoneInfoConvertTime2Example"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 23, 2023
'@LastModified July 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttime?view=netframework-4.8.1#system-timezoneinfo-converttime(system-datetimeoffset-system-timezoneinfo)

Option Explicit

'@Description("The following example converts an array of DateTimeOffset values to times in the Eastern Time zone of the U.S. and Canada.")
'It illustrates that the ConvertTime method takes time zone adjustments into account, because a
'time zone adjustment occurs in both the source and destination time zones at 2:00 A.M. on November 7, 2010.
Public Sub TimeZoneInfoConvertTime2()
Attribute TimeZoneInfoConvertTime2.VB_Description = "The following example converts an array of DateTimeOffset values to times in the Eastern Time zone of the U.S. and Canada."
    TimeZoneInfo.ClearCachedData 'Clear incase timezone was changed.
    
    ' Define times to be converted.
    Dim time1 As DateTime
    Set time1 = DateTime.CreateFromDateTime(2010, 1, 1, 12, 1, 0)
    Dim time2 As DateTime
    Set time2 = DateTime.CreateFromDateTime(2010, 11, 6, 23, 30, 0)
    
    Dim times() As DateTimeOffset
    Objects.ToArray times, _
                    DateTimeOffset.CreateFromDateTime2(time1, TimeZoneInfo.Locale.GetUtcOffset(time1)), _
                    DateTimeOffset.CreateFromDateTime2(time1, TimeSpan.Zero), _
                    DateTimeOffset.CreateFromDateTime2(time2, TimeZoneInfo.Locale.GetUtcOffset(time2)), _
                    DateTimeOffset.CreateFromDateTime2(time2.AddHours(3), TimeZoneInfo.Locale.GetUtcOffset(time2.AddHours(3)))
                    

    
    ' Retrieve the time zone for Eastern Standard Time (U.S. and Canada).
    Dim est As TimeZoneInfo
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
        Dim timeToConvert As DateTimeOffset
        Set timeToConvert = varTimeToConvert
        
        Dim targetTime As DateTimeOffset
        Set targetTime = TimeZoneInfo.ConvertTime2(timeToConvert, est)
        Debug.Print "Converted " & timeToConvert.ToString() & _
                     " to " & targetTime.ToString() & "."
    Next
End Sub

'    The example displays the following output:
'    Local time zone: (UTC-08:00) Pacific Time (US & Canada)
'
'    Converted 1/1/2010 12:01:00 PM -08:00 to 1/1/2010 3:01:00 PM -05:00.
'    Converted 1/1/2010 12:01:00 PM +00:00 to 1/1/2010 7:01:00 AM -05:00.
'    Converted 11/6/2010 11:30:00 PM -07:00 to 11/7/2010 1:30:00 AM -05:00.
'    Converted 11/7/2010 2:30:00 AM -08:00 to 11/7/2010 5:30:00 AM -05:00.
