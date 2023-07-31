Attribute VB_Name = "DateTimeOffsetGreaterThanEg"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Operators")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.op_greaterthan?view=netframework-4.8.1

Option Explicit

'@Description("The GreaterThan method defines the operation of the greater than operator for DateTimeOffset objects.")
' It enables code such as the following:
Public Sub DateTimeOffsetGreaterThan()
Attribute DateTimeOffsetGreaterThan.VB_Description = "The GreaterThan method defines the operation of the greater than operator for DateTimeOffset objects."
    Dim date1 As IDateTimeOffset
    Set date1 = DateTimeOffset.CreateFromDateTimeParts(2007, 6, 3, 14, 45, 0, TimeSpan.Create(-7, 0, 0))
    Dim date2 As IDateTimeOffset
    Set date2 = DateTimeOffset.CreateFromDateTimeParts(2007, 6, 3, 15, 45, 0, TimeSpan.Create(-6, 0, 0))
    Dim date3 As DateTimeOffset
    Set date3 = DateTimeOffset.CreateFromDateTime2(date1.DateTime, TimeSpan.Create(-6, 0, 0))
    
    Debug.Print DateTimeOffset.GreaterThan(date1, date2)   ' Displays False
    Debug.Print DateTimeOffset.GreaterThan(date1, date3)   ' Displays True
End Sub

' Output:
'    False
'    True
