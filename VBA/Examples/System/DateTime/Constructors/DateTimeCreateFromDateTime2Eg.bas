Attribute VB_Name = "DateTimeCreateFromDateTime2Eg"
'@Folder "Examples.System.DateTime.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 21, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-globalization-calendar)

Option Explicit

''
' The following example calls the
' DateTime(Int32, Int32, Int32, Int32, Int32, Int32, Calendar)
' constructor twice to instantiate two DateTime values.
' The first call instantiates a DateTime value by using a PersianCalendar object.
' Because the Persian calendar cannot be designated as the default calendar for a culture,
' displaying a date in the Persian calendar requires individual calls to its
' PersianCalendar.GetMonth, PersianCalendar.GetDayOfMonth, and PersianCalendar.GetYear methods.
' The second call to the constructor instantiates a DateTime value by using a HijriCalendar object.
' The example changes the current culture to Arabic (Syria) and changes the current culture's
' default calendar to the Hijri calendar. Because Hijri is the current culture's default calendar,
' the Console.WriteLine method uses it to format the date.
' When the previous current culture (which is English (United States) in this case) is restored,
' the Console.WriteLine method uses the current culture's default Gregorian calendar to format the date.
''
Public Sub DateTimeCreateFromDateTime2()
    Debug.Print "Using the Persian Calendar:"
    Dim persian As DotNetLib.PersianCalendar
    Set persian = New DotNetLib.PersianCalendar
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime2(1389, 5, 27, 16, 32, 0, persian)
    Debug.Print date1.ToString()
    Debug.Print VBString.Format(VBString.Unescape("{0}/{1}/{2} {3}{6}{4:D2}{6}{5:D2}\n"), _
                        persian.GetMonth(date1), _
                        persian.GetDayOfMonth(date1), _
                        persian.GetYear(date1), _
                        persian.GetHour(date1), _
                        persian.GetMinute(date1), _
                        persian.GetSecond(date1), _
                        DateTimeFormatInfo.CurrentInfo.TimeSeparator)

    Debug.Print "Using the Hijri Calendar:"
    ' Get current culture so it can later be restored.
    Dim dftCulture As DotNetLib.CultureInfo
    Set dftCulture = CultureInfo.CurrentCulture

    ' Define Hijri calendar.
    Dim hijri As DotNetLib.HijriCalendar
    Set hijri = New DotNetLib.HijriCalendar
    
    ' Make ar-SY the current culture and Hijri the current calendar
    Set CultureInfo.CurrentCulture = CultureInfo.CreateFromName("ar-SY")
    
    Dim current As DotNetLib.CultureInfo
    Set current = CultureInfo.CurrentCulture
    Set current.DateTimeFormat.Calendar = hijri
    Dim dFormat As String
    dFormat = current.DateTimeFormat.ShortDatePattern
    ' Ensure year is displayed as four digits.
    dFormat = Regex.Replace(dFormat, "/yy$", "/yyyy")
    current.DateTimeFormat.ShortDatePattern = dFormat
    Dim date2 As DotNetLib.DateTime
    Set date2 = DateTime.CreateFromDateTime2(1431, 9, 9, 16, 32, 18, hijri)
    Debug.Print VBString.Format("{0} culture using the {1} calendar: {2:g}", current, _
                        GetCalendarName(hijri), date2)
    
    ' Restore previous culture.
    Set CultureInfo.CurrentCulture = dftCulture
    Debug.Print VBString.Format("{0} culture using the {1} calendar: {2:g}", _
                      CultureInfo.CurrentCulture, _
                      GetCalendarName(CultureInfo.CurrentCulture.Calendar), _
                      date2)
End Sub

Private Function GetCalendarName(ByVal cal As DotNetLib.Calendar) As String
    GetCalendarName = Regex.Match(cal.ToString(), "\.(\w+)Calendar").Groups.Item(1).value
End Function

' The example displays the following output:
'       Using the Persian Calendar:
'       8/18/2010 4:32:00 PM
'       5/27/1389 16:32:00
'
'       Using the Hijri Calendar:
'       ar-SY culture using the Hijri calendar: 09/09/1431 04:32 ?
'       en-US culture using the Gregorian calendar: 8/18/2010 4:32 PM


