Attribute VB_Name = "DateTimeTimeOfDayExample"
'Rubberduck annotations
'@Folder "VBADotNetLib.Examples.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.timeofday?view=netframework-4.8.1#examples

'@Notes
' https://learn.microsoft.com/en-us/dotnet/standard/base-types/standard-timespan-format-strings
' https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-timespan-format-strings

Option Explicit

Public Sub DateTimeTimeOfDay()
    Dim dates(3) As IDateTime
    
    Set dates(0) = DateTime.Now
    Set dates(1) = DateTime.CreateFromDateTime(2013, 9, 14, 9, 28, 0)
    Set dates(2) = DateTime.CreateFromDateTime(2011, 5, 28, 10, 35, 0)
    Set dates(3) = DateTime.CreateFromDateTime(1979, 12, 25, 14, 30, 0)
    
    Dim varDateTime As Variant
    For Each varDateTime In dates    ' Iterate through each element.
        Dim dtObject As IDateTime
        Set dtObject = varDateTime
        Debug.Print "Day: " & dtObject.DateOnly.ToString2("d") & " Time: " & dtObject.TimeOfDay.ToString2("g")
        Debug.Print "Day: " & dtObject.ToString2("d") & " Time: " & dtObject.ToString2("t")
    Next
End Sub

' The example displays output like the following:
'    Day: 7/25/2012 Time: 10:08:12.9713744
'    Day: 7/25/2012 Time: 10:08 AM
'
'    Day: 9/14/2013 Time: 9:28:00
'    Day: 9/14/2013 Time: 9:28 AM
'
'    Day: 5/28/2011 Time: 10:35:00
'    Day: 5/28/2011 Time: 10:35 AM
'
'    Day: 12/25/1979 Time: 14:30:00
'    Day: 12/25/1979 Time: 2:30 PM
