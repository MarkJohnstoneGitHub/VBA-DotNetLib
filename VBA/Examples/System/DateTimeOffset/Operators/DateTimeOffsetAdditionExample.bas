Attribute VB_Name = "DateTimeOffsetAdditionExample"
'@Folder "Examples.System.DateTimeOffset.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.op_addition?view=netframework-4.8.1

Option Explicit

''
' The Addition method defines the addition operation for DateTimeOffset values.
''
Public Sub DateTimeOffsetAddition()
    Dim date1 As DotNetLib.DateTimeOffset
    Set date1 = DateTimeOffset.CreateFromDateTimeParts(2008, 1, 1, 13, 32, 45, TimeSpan.Create(-5, 0, 0))
    Dim interval1 As DotNetLib.TimeSpan
    Set interval1 = TimeSpan.Create2(202, 3, 30, 0)
    Dim interval2 As DotNetLib.TimeSpan
    Set interval2 = TimeSpan.Create2(5, 0, 0, 0)
    Dim date2 As DotNetLib.DateTimeOffset
    
    Debug.Print date1.ToString()                ' Displays 1/1/2008 1:32:45 PM -05:00
    Set date2 = DateTimeOffset.Addition(date1, interval1)
    Debug.Print date2.ToString()                ' Displays 7/21/2008 5:02:45 PM -05:00
    
    Set date2 = DateTimeOffset.Addition(date2, interval2)
    Debug.Print date2.ToString()                ' Displays 7/26/2008 5:02:45 PM -05:00
End Sub

'Output:
'    1/01/2008 1:32:45 PM -05:00
'    21/07/2008 5:02:45 PM -05:00
'    26/07/2008 5:02:45 PM -05:00
