Attribute VB_Name = "DateTimeTodayExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.today?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the Date property to retrieve the current date.")
' It also illustrates how a DateTime value can be formatted using some of the
' standard date and time format strings. Note that the output produced by the
' third call to the ToString(String) method uses the g format specifier to
' include the time component, which is zero.
Public Sub DateTimeToday()
Attribute DateTimeToday.VB_Description = "The following example uses the Date property to retrieve the current date."
    ' Get the current date.
    Dim thisDay As IDateTime
    Set thisDay = DateTime.Today
    ' Display the date in the default (general) format.
    Debug.Print thisDay.ToString()
    Debug.Print
    ' Display the date in a variety of formats.
    Debug.Print thisDay.ToString2("d")
    Debug.Print thisDay.ToString2("D")
    Debug.Print thisDay.ToString2("g")
End Sub

' The example displays output similar to the following:
'    5/3/2012 12:00:00 AM
'
'    5/3/2012
'    Thursday, May 03, 2012
'    5/3/2012 12:00 AM
