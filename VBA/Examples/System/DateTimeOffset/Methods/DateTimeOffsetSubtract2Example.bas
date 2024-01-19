Attribute VB_Name = "DateTimeOffsetSubtract2Example"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.subtract?view=netframework-4.8.1#system-datetimeoffset-subtract(system-timespan)

Option Explicit

''
' The following example illustrates subtraction that uses the Subtract method.
''
Public Sub DateTimeOffsetSubtract2()
    Dim offsetDate As DotNetLib.DateTimeOffset
    Set offsetDate = DateTimeOffset.CreateFromDateTimeParts(2007, 12, 3, 11, 30, 0, TimeSpan.Create(-8, 0, 0))
    Dim pvtDuration As DotNetLib.TimeSpan
    Set pvtDuration = TimeSpan.Create2(7, 18, 0, 0)
    Debug.Print offsetDate.Subtract2(pvtDuration).ToString()
End Sub

' Output:
' 11/25/2007 5:30:00 PM -08:00
