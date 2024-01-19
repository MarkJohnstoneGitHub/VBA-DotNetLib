Attribute VB_Name = "DateTimeOffsetParse3Example"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 25, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.parse?view=netframework-4.8.1#system-datetimeoffset-parse(system-string-system-iformatprovider-system-globalization-datetimestyles)

Option Explicit

''
' The following example illustrates the effect of passing the DateTimeStyles.AssumeLocal,
' DateTimeStyles.AssumeUniversal, and DateTimeStyles.AdjustToUniversal values to the
' styles parameter of the Parse(String, IFormatProvider, DateTimeStyles) method.
''
Public Sub DateTimeOffsetParse3()
    Dim dateString As String
    Dim offsetDate As DotNetLib.DateTimeOffset
    
    dateString = "05/01/2008 6:00:00"
    ' Assume time is local
    Set offsetDate = DateTimeOffset.Parse3(dateString, Nothing, DateTimeStyles.DateTimeStyles_AssumeLocal)
    Debug.Print offsetDate.ToString() ' Displays 5/1/2008 6:00:00 AM -07:00

    ' Assume time is UTC
    Set offsetDate = DateTimeOffset.Parse3(dateString, Nothing, DateTimeStyles.DateTimeStyles_AssumeUniversal)
    Debug.Print offsetDate.ToString()  ' Displays 5/1/2008 6:00:00 AM +00:00

    ' Parse and convert to UTC
    dateString = "05/01/2008 6:00:00AM +5:00"
    Set offsetDate = DateTimeOffset.Parse3(dateString, Nothing, DateTimeStyles.DateTimeStyles_AdjustToUniversal)
    Debug.Print offsetDate.ToString()  ' Displays 5/1/2008 1:00:00 AM +00:00
End Sub

