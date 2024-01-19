Attribute VB_Name = "DateTimeOffsetLessThanExample"
'@Folder "Examples.System.DateTimeOffset.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.op_lessthan?view=netframework-4.8.1#remarks

Option Explicit

''
' The LessThan method defines the operation of the less than operator for
' DateTimeOffset objects.
''
Public Sub DateTimeOffsetLessThan()
    Dim date1 As DotNetLib.DateTimeOffset
    Set date1 = DateTimeOffset.CreateFromDateTimeParts(2007, 6, 3, 14, 45, 0, TimeSpan.Create(-7, 0, 0))
    Dim date2 As DotNetLib.DateTimeOffset
    Set date2 = DateTimeOffset.CreateFromDateTimeParts(2007, 6, 3, 15, 45, 0, TimeSpan.Create(-6, 0, 0))
    Dim date3 As DotNetLib.DateTimeOffset
    Set date3 = DateTimeOffset.CreateFromDateTime2(date1.DateTime, TimeSpan.Create(-8, 0, 0))
    
    Debug.Print DateTimeOffset.LessThan(date1, date2)     ' Displays False
    Debug.Print DateTimeOffset.LessThan(date1, date3)     ' Displays True
End Sub

'Output:
'    False
'    True
