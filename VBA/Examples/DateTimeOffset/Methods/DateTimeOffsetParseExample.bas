Attribute VB_Name = "DateTimeOffsetParseExample"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.parse?view=netframework-4.8.1#system-datetimeoffset-parse(system-string)

Option Explicit

'@Description("The following example calls the Parse(String) method to parse several date and time strings.")
Public Sub DateTimeOffsetParse()
Attribute DateTimeOffsetParse.VB_Description = "The following example calls the Parse(String) method to parse several date and time strings."
    Dim dateString As String
    Dim offsetDate As IDateTimeOffset
    
    ' String with date only
    dateString = "05/01/2008"
    Set offsetDate = DateTimeOffset.Parse(dateString)
    Debug.Print offsetDate.ToString()

    ' String with time only
    dateString = "11:36 PM"
    Set offsetDate = DateTimeOffset.Parse(dateString)
    Debug.Print offsetDate.ToString()
    
    ' String with date and offset
    dateString = "05/01/2008 +1:00"
    Set offsetDate = DateTimeOffset.Parse(dateString)
    Debug.Print offsetDate.ToString()

    ' String with day abbreviation
    dateString = "Thu May 01, 2008"
    Set offsetDate = DateTimeOffset.Parse(dateString)
    Debug.Print offsetDate.ToString()
End Sub

' Output:
'    5/1/2008 12:00:00 AM -07:00
'    7/20/2023 11:36:00 PM -07:00
'    5/1/2008 12:00:00 AM +01:00
'    5/1/2008 12:00:00 AM -07:00
