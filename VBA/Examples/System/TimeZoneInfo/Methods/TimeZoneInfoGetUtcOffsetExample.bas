Attribute VB_Name = "TimeZoneInfoGetUtcOffsetExample"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 26, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.getutcoffset?view=netframework-4.8.1#system-timezoneinfo-getutcoffset(system-datetime)

Option Explicit

''
' The following example illustrates the use of the GetUtcOffset(DateTime)
' method with different time zones and with date values that have different
' Kind property values.
''
Public Sub TimeZoneInfoGetUtcOffset()
    Dim cst As DotNetLib.TimeZoneInfo
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

Private Sub ShowOffset(ByVal time As DotNetLib.DateTime, ByVal timeZone As DotNetLib.TimeZoneInfo)
    Dim convertedTime As DotNetLib.DateTime
    Set convertedTime = time
    Dim pvtOffset As DotNetLib.TimeSpan
    
    If (time.Kind = DateTimeKind.DateTimeKind_Local And Not timeZone.Equals(TimeZoneInfo.Locale)) Then
        Set convertedTime = TimeZoneInfo.ConvertTime3(time, TimeZoneInfo.Locale, timeZone)
    ElseIf (time.Kind = DateTimeKind.DateTimeKind_Utc And Not timeZone.Equals(TimeZoneInfo.Utc)) Then
        Set convertedTime = TimeZoneInfo.ConvertTime3(time, TimeZoneInfo.Utc, timeZone)
    End If
    Set pvtOffset = timeZone.GetUtcOffset(time)
    If DateTime.Equality(time, convertedTime) Then
        Debug.Print VBString.Format("{0} {1} ", time, _
                        IIf(timeZone.IsDaylightSavingTime(time), timeZone.DaylightName, timeZone.StandardName))
        Debug.Print VBString.Format("   It differs from UTC by {0} hours, {1} minutes.", _
                        pvtOffset.Hours, _
                        pvtOffset.Minutes)
    Else
        Debug.Print VBString.Format("{0} {1} ", time, _
                        IIf(time.Kind = DateTimeKind.DateTimeKind_Utc, "UTC", TimeZoneInfo.Locale.Id))
        Debug.Print VBString.Format("   converts to {0} {1}.", _
                          convertedTime, _
                          timeZone.Id)
        Debug.Print VBString.Format("   It differs from UTC by {0} hours, {1} minutes.", _
                          pvtOffset.Hours, pvtOffset.Minutes)
    End If
    Debug.Print
End Sub

' The example produces the following output:
'
'       6/12/2006 11:00:00 AM Pacific Daylight Time
'          It differs from UTC by -7 hours, 0 minutes.
'
'       11/4/2007 1:00:00 AM Pacific Standard Time
'          It differs from UTC by -8 hours, 0 minutes.
'
'       12/10/2006 3:00:00 PM Pacific Standard Time
'          It differs from UTC by -8 hours, 0 minutes.
'
'       3/11/2007 2:30:00 AM Pacific Standard Time
'          It differs from UTC by -8 hours, 0 minutes.
'
'       2/2/2007 8:35:46 PM UTC
'          converts to 2/2/2007 12:35:46 PM Pacific Standard Time.
'          It differs from UTC by -8 hours, 0 minutes.
'
'       6/12/2006 11:00:00 AM UTC
'          It differs from UTC by 0 hours, 0 minutes.
'
'       11/4/2007 1:00:00 AM UTC
'          It differs from UTC by 0 hours, 0 minutes.
'
'       12/10/2006 3:00:00 AM UTC
'          It differs from UTC by 0 hours, 0 minutes.
'
'       3/11/2007 2:30:00 AM UTC
'          It differs from UTC by 0 hours, 0 minutes.
'
'       2/2/2007 12:35:46 PM Pacific Standard Time
'          converts to 2/2/2007 8:35:46 PM UTC.
'          It differs from UTC by 0 hours, 0 minutes.
'
'       6/12/2006 11:00:00 AM Central Daylight Time
'          It differs from UTC by -5 hours, 0 minutes.
'
'       11/4/2007 1:00:00 AM Central Standard Time
'          It differs from UTC by -6 hours, 0 minutes.
'
'       12/10/2006 3:00:00 PM Central Standard Time
'          It differs from UTC by -6 hours, 0 minutes.
'
'       3/11/2007 2:30:00 AM Central Standard Time
'          It differs from UTC by -6 hours, 0 minutes.
'
'       11/14/2007 12:00:00 AM Pacific Standard Time
'          converts to 11/14/2007 2:00:00 AM Central Standard Time.
'          It differs from UTC by -6 hours, 0 minutes.
