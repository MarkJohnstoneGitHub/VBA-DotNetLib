Attribute VB_Name = "DateTimeTicksExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.ticks?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the Ticks property to display the number of ticks
' that have elapsed since the beginning of the twenty-first century and to
' instantiate a TimeSpan object. The TimeSpan object is then used to display
' the elapsed time using several other time intervals.
''
Public Sub DateTimeTicks()
    Dim centuryBegin As DotNetLib.DateTime
    Set centuryBegin = DateTime.CreateFromDate(2001, 1, 1)
    Dim currentDate As DotNetLib.DateTime
    Set currentDate = DateTime.Now
    
    Dim ElapsedTicks As LongLong
    ElapsedTicks = currentDate.Ticks - centuryBegin.Ticks
    
    Dim elapsedSpan As DotNetLib.TimeSpan
    Set elapsedSpan = TimeSpan.CreateFromTicks(ElapsedTicks)
    Debug.Print VBString.Format("Elapsed from the beginning of the century to {0:f}:", _
                       currentDate)
    Debug.Print VBString.Format("   {0:N0} nanoseconds", ElapsedTicks * 100)
    Debug.Print VBString.Format("   {0:N0} ticks", ElapsedTicks)
    Debug.Print VBString.Format("   {0:N2} seconds", elapsedSpan.totalSeconds)
    Debug.Print VBString.Format("   {0:N2} minutes", elapsedSpan.TotalMinutes)
    Debug.Print VBString.Format("   {0:N0} days, {1} hours, {2} minutes, {3} seconds", _
                            elapsedSpan.Days, elapsedSpan.Hours, _
                            elapsedSpan.Minutes, elapsedSpan.Seconds)
End Sub

' This example displays an output similar to the following:
'
' Elapsed from the beginning of the century to Thursday, 14 November 2019 18:21:
'    595,448,498,171,000,000 nanoseconds
'    5,954,484,981,710,000 ticks
'    595,448,498.17 seconds
'    9,924,141.64 minutes
'    6,891 days, 18 hours, 21 minutes, 38 seconds

