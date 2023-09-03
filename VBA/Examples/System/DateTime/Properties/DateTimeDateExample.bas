Attribute VB_Name = "DateTimeDateExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified September 3, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.date?view=netframework-4.8.1#examples

Option Explicit

' The following example uses the Date property to extract the date component of
' a DateTime value with its time component set to zero (or 0:00:00, or midnight).
' It also illustrates that, depending on the format string used when displaying
' the DateTime value, the time component can continue to appear in formatted output.
'
Public Sub DateTimeDate()
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTime(2008, 6, 1, 7, 47, 0)
    Debug.Print date1.ToString
    
    ' Get date-only portion of date, without its time.
    Dim pvtDateOnly As IDateTime
    Set pvtDateOnly = date1.Date()
    
    ' Display date using short date string.
    Debug.Print pvtDateOnly.ToString2("d")
    ' Display date using 24-hour clock.
    Debug.Print pvtDateOnly.ToString2("g")
    Debug.Print pvtDateOnly.ToString2("MM/dd/yyyy HH:mm")
End Sub

' The example displays output like the following output:
'       6/1/2008 7:47:00 AM
'       6/1/2008
'       6/1/2008 12:00 AM
'       06/01/2008 00:00
