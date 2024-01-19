Attribute VB_Name = "DateTimeOffsetDayOfWeekExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.dayofweek?view=netframework-4.8.1#examples

Option Explicit

''
' The following example displays the weekday name of the first day of each month
' of the year 2008.
''
Public Sub DateTimeOffsetDayOfWeek()
    Dim startOfMonth  As DotNetLib.DateTimeOffset
    Set startOfMonth = DateTimeOffset.CreateFromDateTimeParts(2008, 1, 1, 0, 0, 0, DateTimeOffset.Now.Offset)
    Dim pvtYear As Long
    pvtYear = startOfMonth.Year
    Do
        Debug.Print VBString.Format("{0:MMM d, yyyy} is a {1}.", startOfMonth, DayOfWeekHelper.ToString(startOfMonth.DayOfWeek))
        Set startOfMonth = startOfMonth.AddMonths(1)
    Loop While startOfMonth.Year = pvtYear
End Sub

' This example writes the following output to the console:
'    Jan 1, 2008 is a Tuesday.
'    Feb 1, 2008 is a Friday.
'    Mar 1, 2008 is a Saturday.
'    Apr 1, 2008 is a Tuesday.
'    May 1, 2008 is a Thursday.
'    Jun 1, 2008 is a Sunday.
'    Jul 1, 2008 is a Tuesday.
'    Aug 1, 2008 is a Friday.
'    Sep 1, 2008 is a Monday.
'    Oct 1, 2008 is a Wednesday.
'    Nov 1, 2008 is a Saturday.
'    Dec 1, 2008 is a Monday.
