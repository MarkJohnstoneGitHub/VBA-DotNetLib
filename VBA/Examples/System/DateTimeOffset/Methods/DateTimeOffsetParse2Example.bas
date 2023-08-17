Attribute VB_Name = "DateTimeOffsetParse2Example"
'@Folder("Examples.System.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August, 17 2023
'@LastModified August, 17 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.parse?view=netframework-4.8.1#system-datetimeoffset-parse(system-string-system-iformatprovider)

Option Explicit

' The following example parses date and time strings that are formatted for the
' fr-fr culture and displays them using the local system's default en-us culture.
Public Sub DateTimeOffsetParse2()
    Dim fmt As IFormatProvider 'DateTimeFormatInfo
    Set fmt = CultureInfo.Create2("fr-fr").DateTimeFormat
    Dim dateString As String
    Dim offsetDate As IDateTimeOffset
    
    dateString = "03-12-07"
    Set offsetDate = DateTimeOffset.Parse2(dateString, fmt)
    Debug.Print dateString; " returns "; offsetDate.ToString()

    dateString = "15/09/07 08:45:00 +1:00"
    Set offsetDate = DateTimeOffset.Parse2(dateString, fmt)
    Debug.Print dateString; " returns "; offsetDate.ToString()
    
    dateString = "mar. 1 janvier 2008 1:00:00 +1:00"
    Set offsetDate = DateTimeOffset.Parse2(dateString, fmt)
    Debug.Print dateString; " returns "; offsetDate.ToString()
End Sub

' The example displays the following output to the console:
'    03-12-07 returns 12/3/2007 12:00:00 AM -08:00
'    15/09/07 08:45:00 +1:00 returns 9/15/2007 8:45:00 AM +01:00
'    mar. 1 janvier 2008 1:00:00 +1:00 returns 1/1/2008 1:00:00 AM +01:00


