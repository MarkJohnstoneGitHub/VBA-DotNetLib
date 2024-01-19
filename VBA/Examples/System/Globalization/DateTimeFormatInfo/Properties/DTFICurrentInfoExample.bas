Attribute VB_Name = "DTFICurrentInfoExample"
'@Folder "Examples.System.Globalization.DateTimeFormatInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.currentinfo?view=netframework-4.8.1#examples

Option Explicit

' The following example uses the CurrentInfo property to retrieve a DateTimeFormatInfo
' object that represents the formatting conventions of the current culture, which in
' this case is the en-US culture. It then displays the format string and the result
' string for six formatting properties.
Public Sub DateTimeFormatInfoCurrentInfo()
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime(2016, 5, 28, 10, 28, 0)
    Dim dtfi As DotNetLib.DateTimeFormatInfo
    Set dtfi = DateTimeFormatInfo.CurrentInfo
    
    Debug.Print VBString.Format("Date and Time Formats for {0:u} in the {1} Culture:", _
                                date1, CultureInfo.CurrentCulture.Name); VBA.vbNewLine
    Debug.Print VBString.Format("{0,-22} {1,-20} {2,-30}", "Long Date Pattern", _
                                dtfi.LongDatePattern, _
                                date1.ToString2(dtfi.LongDatePattern))
    Debug.Print VBString.Format("{0,-22} {1,-20} {2,-30}", "Long Time Pattern", _
                                dtfi.LongTimePattern, _
                                date1.ToString2(dtfi.LongTimePattern))
    Debug.Print VBString.Format("{0,-22} {1,-20} {2,-30}", "Month/Day Pattern", _
                                dtfi.MonthDayPattern, _
                                date1.ToString2(dtfi.MonthDayPattern))
    
    Debug.Print VBString.Format("{0,-22} {1,-20} {2,-30}", "Short Date Pattern", _
                                dtfi.ShortDatePattern, _
                                date1.ToString2(dtfi.ShortDatePattern))

    Debug.Print VBString.Format("{0,-22} {1,-20} {2,-30}", "Short Time Pattern", _
                                dtfi.ShortTimePattern, _
                                date1.ToString2(dtfi.ShortTimePattern))

    Debug.Print VBString.Format("{0,-22} {1,-20} {2,-30}", "Year/Month Pattern", _
                                dtfi.YearMonthPattern, _
                                date1.ToString2(dtfi.YearMonthPattern))
End Sub

' The example displays the following output:
'    Date and Time Formats for 2016-05-28 10:28:00Z in the en-US Culture:
'
'    Long Date Pattern      dddd, MMMM d, yyyy   Saturday, May 28, 2016
'    Long Time Pattern      h:mm:ss tt           10:28:00 AM
'    Month/Day Pattern      MMMM d               May 28
'    Short Date Pattern     M/d/yyyy             5/28/2016
'    Short Time Pattern     h:mm tt              10:28 AM
'    Year/Month Pattern     MMMM yyyy            May 2016


