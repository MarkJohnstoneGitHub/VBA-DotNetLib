Attribute VB_Name = "DTFICurrentInfoExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 1, 2023

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
    
    Debug.Print "Date and Time Formats for "; date1.ToString2("u"); _
                " in the "; CultureInfo.CurrentCulture.Name; " Culture:"

    Debug.Print Align("Long Date Pattern", 22, Justify_Left); " "; Align(dtfi.LongDatePattern, 20, Justify_Left);
    Debug.Print date1.ToString2(dtfi.LongDatePattern)
   
    Debug.Print Align("Long Time Pattern", 22, Justify_Left); " "; Align(dtfi.LongTimePattern, 20, Justify_Left);
    Debug.Print date1.ToString2(dtfi.LongTimePattern)

    Debug.Print Align("Month/Day Pattern", 22, Justify_Left); " "; Align(dtfi.MonthDayPattern, 20, Justify_Left);
    Debug.Print date1.ToString2(dtfi.MonthDayPattern)

    Debug.Print Align("Short Date Pattern", 22, Justify_Left); " "; Align(dtfi.ShortDatePattern, 20, Justify_Left);
    Debug.Print date1.ToString2(dtfi.ShortDatePattern)
    
    Debug.Print Align("Short Time Pattern", 22, Justify_Left); " "; Align(dtfi.ShortTimePattern, 20, Justify_Left);
    Debug.Print date1.ToString2(dtfi.ShortTimePattern)
    
    Debug.Print Align("YearMonthPattern", 22, Justify_Left); " "; Align(dtfi.YearMonthPattern, 20, Justify_Left);
    Debug.Print date1.ToString2(dtfi.YearMonthPattern)
    
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
