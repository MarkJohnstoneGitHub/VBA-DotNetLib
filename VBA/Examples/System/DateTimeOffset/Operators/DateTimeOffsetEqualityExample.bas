Attribute VB_Name = "DateTimeOffsetEqualityExample"
'@Folder "Examples.System.DateTimeOffset.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.op_equality?view=netframework-4.8.1

Option Explicit

''
' The Equality method defines the operation of the equality operator for
' DateTimeOffset objects.
''
Public Sub DateTimeOffsetEquality()
    Dim date1 As DotNetLib.DateTimeOffset
    Set date1 = DateTimeOffset.CreateFromDateTimeParts(2007, 6, 3, 14, 45, 0, TimeSpan.Create(-7, 0, 0))
    Dim date2 As DotNetLib.DateTimeOffset
    Set date2 = DateTimeOffset.CreateFromDateTimeParts(2007, 6, 3, 15, 45, 0, TimeSpan.Create(-6, 0, 0))
    Dim date3 As DotNetLib.DateTimeOffset
    Set date3 = DateTimeOffset.CreateFromDateTime2(date1.DateTime, TimeSpan.Create(-6, 0, 0))
    
    Debug.Print DateTimeOffset.Equality(date1, date2)   ' Displays True
    Debug.Print DateTimeOffset.Equality(date1, date3)   ' Displays False
End Sub

' Output:
'    True
'    False
