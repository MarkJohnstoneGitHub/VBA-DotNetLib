Attribute VB_Name = "DateTimeTicksExample"
'@Folder("Examples.DateTime")

' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.ticks?view=netframework-4.8.1#examples

Option Explicit


' The following example uses the Ticks property to display the number of ticks
' that have elapsed since the beginning of the twenty-first century and to
' instantiate a TimeSpan object. The TimeSpan object is then used to display
' the elapsed time using several other time intervals.
Public Sub DateTimeTicks()
    Dim centuryBegin As DateTime
    Set centuryBegin = DateTime.CreateFromDate(2001, 1, 1)
    Dim currentDate As DateTime
    Set currentDate = DateTime.Now
    
    Dim elapsedTicks As LongLong
    elapsedTicks = currentDate.Ticks - centuryBegin.Ticks
    
    Dim elapsedSpan As DotNetLib.TimeSpan
    With New DotNetLib.TimeSpan
        Set elapsedSpan = .CreateFromTicks(elapsedTicks)
    End With
    Debug.Print "Elapsed from the beginning of the century to " & currentDate.ToString2("f")
    Debug.Print "   " & elapsedTicks * 100 & " nanoseconds"
    Debug.Print "   " & elapsedTicks & " ticks"
    Debug.Print "   " & elapsedSpan.TotalSeconds & " seconds"
    Debug.Print "   " & elapsedSpan.TotalMinutes & " minutes"
    Debug.Print "   " & elapsedSpan.days & " days, " & elapsedSpan.hours & " hours, " & elapsedSpan.minutes & " minutes, " & elapsedSpan.seconds & " seconds"
    
' This example displays an output similar to the following:
'
' Elapsed from the beginning of the century to Thursday, 14 November 2019 18:21:
'    595,448,498,171,000,000 nanoseconds
'    5,954,484,981,710,000 ticks
'    595,448,498.17 seconds
'    9,924,141.64 minutes
'    6,891 days, 18 hours, 21 minutes, 38 seconds
End Sub


