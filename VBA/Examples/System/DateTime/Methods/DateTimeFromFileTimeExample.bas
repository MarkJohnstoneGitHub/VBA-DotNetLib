Attribute VB_Name = "DateTimeFromFileTimeExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 11, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.fromfiletime?view=netframework-4.8.1#notes-to-callers

Option Explicit

''
' For example, the transition from standard time to daylight saving time occurs
' in the U.S. Pacific Time zone on March 14, 2010, at 2:00 A.M., when the time
' advances by one hour, to 3:00 A.M. This hour interval is an invalid time,
' that is, a time interval that does not exist in this time zone.
' The following example shows that when a time that falls within this range is
' converted to a long integer value by the ToFileTime() method and is then
' restored by the FromFileTime(Int64) method, the original value is adjusted
' to become a valid time. You can determine whether a particular date and time
' value may be subject to modification by passing it to the
' IsInvalidTime(DateTime) method, as the example illustrates.
''
Public Sub DateTimeFromFileTime()
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime(2023, 10, 1, 2, 30, 0)
    Debug.Print VBString.Format("Invalid Time: {0}", _
                        TimeZoneInfo.Locale.IsInvalidTime(date1))
    Dim ft As LongLong
    ft = date1.ToFileTime()
    Dim date2 As DotNetLib.DateTime
    Set date2 = DateTime.FromFileTime(ft)
    Debug.Print VBString.Format("{0} -> {1}", date1, date2)
End Sub

' The example displays the following output for local time zone of AUS Eastern Standard Time:
'       Invalid Time: True
'       1/10/2023 2:30:00 AM -> 1/10/2023 3:30:00 AM

' The example displays the following output for local time zone of US Pacific Standard Time:
'    Invalid Time: False
'    10/1/2023 2:30:00 AM -> 10/1/2023 2:30:00 AM


