Attribute VB_Name = "DateTimeTicksExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.ticks?view=netframework-4.8.1#examples

Option Explicit


'@Description("The following example uses the Ticks property to display the number of ticks that have elapsed since the beginning of the twenty-first century and to instantiate a TimeSpan object.")
' The TimeSpan object is then used to display
' the elapsed time using several other time intervals.
Public Sub DateTimeTicks()
Attribute DateTimeTicks.VB_Description = "The following example uses the Ticks property to display the number of ticks that have elapsed since the beginning of the twenty-first century and to instantiate a TimeSpan object."
    Dim centuryBegin As IDateTime
    Set centuryBegin = DateTime.CreateFromDate(2001, 1, 1)
    Dim currentDate As IDateTime
    Set currentDate = DateTime.Now
    
    Dim elapsedTicks As LongLong
    elapsedTicks = currentDate.Ticks - centuryBegin.Ticks
    
    Dim elapsedSpan As ITimeSpan
    Set elapsedSpan = TimeSpan.CreateFromTicks(elapsedTicks)
    Debug.Print "Elapsed from the beginning of the century to " & currentDate.ToString2("f")
    Debug.Print "   " & VBA.format$(elapsedTicks * 100, "#,##0") & " nanoseconds"
    Debug.Print "   " & VBA.format$(elapsedTicks, "#,##0") & " ticks"
    Debug.Print "   " & VBA.format$(elapsedSpan.totalSeconds, "#,##0.00") & " seconds"
    Debug.Print "   " & VBA.format$(elapsedSpan.TotalMinutes, "#,##0.00") & " minutes"
    Debug.Print "   " & elapsedSpan.days & " days, " & elapsedSpan.Hours & " hours, " & elapsedSpan.Minutes & " minutes, " & elapsedSpan.Seconds & " seconds"
End Sub

' This example displays an output similar to the following:
'
' Elapsed from the beginning of the century to Thursday, 14 November 2019 18:21:
'    595,448,498,171,000,000 nanoseconds
'    5,954,484,981,710,000 ticks
'    595,448,498.17 seconds
'    9,924,141.64 minutes
'    6,891 days, 18 hours, 21 minutes, 38 seconds
