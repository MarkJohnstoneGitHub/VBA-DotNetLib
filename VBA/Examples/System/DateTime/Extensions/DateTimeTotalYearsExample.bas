Attribute VB_Name = "DateTimeTotalYearsExample"
'@Folder("Examples.System.DateTime.Extensions")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 3, 2023
'@LastModified November 3, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference
' https://stackoverflow.com/questions/9/how-do-i-calculate-someones-age-based-on-a-datetime-type-birthday
' https://stackoverflow.com/a/1595311/10759363

Option Explicit

''
' Testing DateTime.TotalYears(startDateTime,endDateTime)
' Obtains the total number of years between two datetime objects
''
Public Sub DateTimeTotalYears()
    Dim startDate As DotNetLib.DateTime
    Dim endDate As DotNetLib.DateTime
    Set startDate = DateTime.CreateFromDate(2000, 2, 29)
    Set endDate = DateTime.CreateFromDate(2023, 2, 28)
    Debug.Print VBAString.Format("Number of years between {0} and {1} = {2} years", startDate, endDate, DateTime.TotalYears(startDate, endDate))
    Debug.Print VBAString.Format("Number of years between {0} and {1} = {2} years", endDate, startDate, DateTime.TotalYears(endDate, startDate))
    Debug.Print
    
    Set endDate = DateTime.CreateFromDate(2023, 3, 1)
    Debug.Print VBAString.Format("Number of years between {0} and {1} = {2} years", startDate, endDate, DateTime.TotalYears(startDate, endDate))
    Debug.Print VBAString.Format("Number of years between {0} and {1} = {2} years", endDate, startDate, DateTime.TotalYears(endDate, startDate))
    Debug.Print
    
    Set startDate = DateTime.CreateFromDate(1901, 5, 30)
    Set endDate = DateTime.CreateFromDate(2000, 5, 29)
    Debug.Print VBAString.Format("Number of years between {0} and {1} = {2} years", startDate, endDate, DateTime.TotalYears(startDate, endDate))
    Debug.Print VBAString.Format("Number of years between {0} and {1} = {2} years", endDate, startDate, DateTime.TotalYears(endDate, startDate))
    Debug.Print
    
    Set startDate = DateTime.CreateFromDate(1901, 5, 30)
    Set endDate = DateTime.CreateFromDate(2000, 5, 30)
    Debug.Print VBAString.Format("Number of years between {0} and {1} = {2} years", startDate, endDate, DateTime.TotalYears(startDate, endDate))
    Debug.Print VBAString.Format("Number of years between {0} and {1} = {2} years", endDate, startDate, DateTime.TotalYears(endDate, startDate))
End Sub

'/*
'This code produces the following output.
'
' Number of years between 29/02/2000 12:00:00 AM and 28/02/2023 12:00:00 AM = 22 years
' Number of years between 28/02/2023 12:00:00 AM and 29/02/2000 12:00:00 AM = -22 years
'
' Number of years between 29/02/2000 12:00:00 AM and 1/03/2023 12:00:00 AM = 23 years
' Number of years between 1/03/2023 12:00:00 AM and 29/02/2000 12:00:00 AM = -23 years
'
' Number of years between 30/05/1901 12:00:00 AM and 29/05/2000 12:00:00 AM = 98 years
' Number of years between 29/05/2000 12:00:00 AM and 30/05/1901 12:00:00 AM = -98 years
'
' Number of years between 30/05/1901 12:00:00 AM and 30/05/2000 12:00:00 AM = 99 years
' Number of years between 30/05/2000 12:00:00 AM and 30/05/1901 12:00:00 AM = -99 years
'*/
