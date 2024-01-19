Attribute VB_Name = "DateTimeDayOfYearExample"
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.dayofyear?view=netframework-4.8.1#examples

Option Explicit

''
' The following example displays the day of the year of December 31 for the years
' 2010-2020 in the Gregorian calendar. Note that the example shows that December 31
' is the 366th day of the year in leap years."
''
Public Sub DateTimeDayOfYear()
    Dim dec31 As DotNetLib.DateTime
    Set dec31 = DateTime.CreateFromDate(2010, 12, 31)
    Dim ctr As Long
    For ctr = 0 To 10
        Dim dateToDisplay As DotNetLib.DateTime
        Set dateToDisplay = dec31.AddYears(ctr)
        Debug.Print VBString.Format("{0:d}: day {1} of {2} {3}", dateToDisplay, _
                           dateToDisplay.DayOfYear, _
                           dateToDisplay.Year, _
                           IIf(DateTime.IsLeapYear(dateToDisplay.Year), " (Leap Year)", vbNullString))
    Next ctr
End Sub

' The example displays the following output:
'       12/31/2010: day 365 of 2010
'       12/31/2011: day 365 of 2011
'       12/31/2012: day 366 of 2012 (Leap Year)
'       12/31/2013: day 365 of 2013
'       12/31/2014: day 365 of 2014
'       12/31/2015: day 365 of 2015
'       12/31/2016: day 366 of 2016 (Leap Year)
'       12/31/2017: day 365 of 2017
'       12/31/2018: day 365 of 2018
'       12/31/2019: day 365 of 2019
'       12/31/2020: day 366 of 2020 (Leap Year)

