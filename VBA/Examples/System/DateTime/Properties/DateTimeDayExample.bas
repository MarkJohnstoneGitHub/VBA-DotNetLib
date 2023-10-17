Attribute VB_Name = "DateTimeDayExample"
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.day?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the Day property.")
Public Sub DateTimePropertyDay()
Attribute DateTimePropertyDay.VB_Description = "The following example demonstrates the Day property."
    Dim moment As IDateTime
    Set moment = DateTime.CreateFromDateTime(1999, 1, 13, 3, 57, 32, 11)
    
    ' Year gets 1999.
    Dim pvtYear As Long
    pvtYear = moment.Year
    Debug.Print "Year "; pvtYear
    
    ' Month gets 1 (January).
    Dim pvtMonth As Long
    pvtMonth = moment.Month
    Debug.Print "Month "; pvtMonth
    
    ' Day gets 13.
    Dim pvtDay As Long
    pvtDay = moment.Day
    Debug.Print "Day "; pvtDay
    
    ' Hour gets 3.
    Dim pvtHour As Long
    pvtHour = moment.Hour
    Debug.Print "Hour "; pvtHour
    
    ' Minute gets 57.
    Dim pvtMinute As Long
    pvtMinute = moment.Minute
    Debug.Print "Minute "; pvtMinute
    
    ' Second gets 32.
    Dim pvtSecond As Long
    pvtSecond = moment.SECOND
    Debug.Print "Second "; pvtSecond
    
    ' Millisecond gets 11.
    Dim pvtMillisecond As Long
    pvtMillisecond = moment.Millisecond
    Debug.Print "Millisecond "; pvtMillisecond
End Sub

' Output:
'    Year 1999
'    Month 1
'    Day 13
'    Hour 3
'    Minute 57
'    Second 32
'    Millisecond 11
