Attribute VB_Name = "DateTimeDayOfYearExample"
'Rubberduck annotations
'@Folder "VBADotNetLib.Examples.DateTime.Properties"

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 09, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.dayofyear?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example displays the day of the year of December 31 for the years 2010-2020 in the Gregorian calendar. Note that the example shows that December 31 is the 366th day of the year in leap years.")
Public Sub DateTimeDayOfYear()
Attribute DateTimeDayOfYear.VB_Description = "The following example displays the day of the year of December 31 for the years 2010-2020 in the Gregorian calendar. Note that the example shows that December 31 is the 366th day of the year in leap years."
    Dim dec31 As DateTime
    Set dec31 = DateTime.CreateFromDate(2010, 12, 31)
    
    Dim ctr As Long
    
    For ctr = 0 To 10
        Dim dateToDisplay As DateTime
        Set dateToDisplay = dec31.AddYears(ctr)
        Debug.Print dateToDisplay.ToString & ": day " & dateToDisplay.DayOfYear & " of " & dateToDisplay.Year & IIf(DateTime.IsLeapYear(dateToDisplay.Year), " (Leap Year)", vbNullString)
        
    Next ctr
    
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
End Sub
