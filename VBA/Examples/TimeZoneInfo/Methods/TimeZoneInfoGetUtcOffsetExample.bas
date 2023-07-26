Attribute VB_Name = "TimeZoneInfoGetUtcOffsetExample"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 26, 2023
'@LastModified July 26, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.getutcoffset?view=netframework-4.8.1#system-timezoneinfo-getutcoffset(system-datetime)

Option Explicit

'@Description("The following example illustrates the use of the GetUtcOffset(DateTime) method with different time zones and with date values that have different Kind property values.")
Public Sub TimeZoneInfoGetUtcOffset()
Attribute TimeZoneInfoGetUtcOffset.VB_Description = "The following example illustrates the use of the GetUtcOffset(DateTime) method with different time zones and with date values that have different Kind property values."
    Dim cst As TimeZoneInfo
    Set cst = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
    
    ShowOffset DateTime.CreateFromDateTime(2006, 6, 12, 11, 0, 0), TimeZoneInfo.Locale
    ShowOffset DateTime.CreateFromDateTime(2007, 11, 4, 1, 0, 0), TimeZoneInfo.Locale
    ShowOffset DateTime.CreateFromDateTime(2006, 12, 10, 15, 0, 0), TimeZoneInfo.Locale
    ShowOffset DateTime.CreateFromDateTime(2007, 3, 11, 2, 30, 0), TimeZoneInfo.Locale
    ShowOffset DateTime.UtcNow, TimeZoneInfo.Locale
    ShowOffset DateTime.CreateFromDateTime(2006, 6, 12, 11, 0, 0), TimeZoneInfo.Utc
    ShowOffset DateTime.CreateFromDateTime(2007, 11, 4, 1, 0, 0), TimeZoneInfo.Utc
    ShowOffset DateTime.CreateFromDateTime(2006, 12, 10, 3, 0, 0), TimeZoneInfo.Utc
    ShowOffset DateTime.CreateFromDateTime(2007, 3, 11, 2, 30, 0), TimeZoneInfo.Utc
    ShowOffset DateTime.Now, TimeZoneInfo.Utc
    ShowOffset DateTime.CreateFromDateTime(2006, 6, 12, 11, 0, 0), cst
    ShowOffset DateTime.CreateFromDateTime(2007, 11, 4, 1, 0, 0), cst
    ShowOffset DateTime.CreateFromDateTime(2006, 12, 10, 15, 0, 0), cst
    ShowOffset DateTime.CreateFromDateTime(2007, 3, 11, 2, 30, 0, 0), cst
    ShowOffset DateTime.CreateFromDateTime(2007, 11, 14, 0, 0, 0, DateTimeKind.DateTimeKind_Local), cst
End Sub

Private Sub ShowOffset(ByVal time As DateTime, ByVal timeZone As TimeZoneInfo)
    Dim convertedTime As DateTime
    Set convertedTime = time
    Dim Offset As TimeSpan
    
    If (time.Kind = DateTimeKind.DateTimeKind_Local And Not timeZone.Equals(TimeZoneInfo.Locale)) Then
        Set convertedTime = TimeZoneInfo.ConvertTime3(time, TimeZoneInfo.Locale, timeZone)
    ElseIf (time.Kind = DateTimeKind.DateTimeKind_Utc And Not timeZone.Equals(TimeZoneInfo.Utc)) Then
        Set convertedTime = TimeZoneInfo.ConvertTime3(time, TimeZoneInfo.Utc, timeZone)
    End If
    Set Offset = timeZone.GetUtcOffset(time)
    If DateTime.Equality(time, convertedTime) Then
        Debug.Print time.ToString; " " & _
                    IIf(timeZone.IsDaylightSavingTime(time), timeZone.DaylightName, timeZone.StandardName)
        Debug.Print "   It differs from UTC by " & Offset.Hours & " hours, " & _
                    Offset.Minutes & " minutes."
    Else
        Debug.Print time.ToString() & " " & _
                    IIf(time.Kind = DateTimeKind.DateTimeKind_Utc, "UTC", TimeZoneInfo.Locale.Id)

        Debug.Print "   converts to " & convertedTime.ToString() & " " & _
                    timeZone.Id & "."
        Debug.Print "   It differs from UTC by " & Offset.Hours & " hours, " & _
                    Offset.Minutes & " minutes."
    End If
    Debug.Print
End Sub
