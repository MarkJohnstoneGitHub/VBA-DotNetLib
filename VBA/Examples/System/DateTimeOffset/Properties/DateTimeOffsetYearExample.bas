Attribute VB_Name = "DateTimeOffsetYearExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 20, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.year?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example displays the year component of a DateTimeOffset value in four different ways:")
'  By retrieving the value of the Year property.
'  By calling the ToString(String) method with the "y" format specifier.
'  By calling the ToString(String) method with the "yy" format specifier.
'  By calling the ToString(String) method with the "yyyy" format specifier.
Public Sub DateTimeOffsetYear()
Attribute DateTimeOffsetYear.VB_Description = "The following example displays the year component of a DateTimeOffset value in four different ways:"
    Dim theTime As IDateTimeOffset
    Set theTime = DateTimeOffset.CreateFromDateTimeParts(2008, 2, 17, 9, 0, 0, DateTimeOffset.Now.Offset)
    
    Debug.Print "The year component of " & theTime.ToString() & " is " & theTime.Year & "."
    Debug.Print "The year component of " & theTime.ToString() & " is" & theTime.ToString2(" y") & "."
    Debug.Print "The year component of " & theTime.ToString() & " is " & theTime.ToString2("yy") & "."
    Debug.Print "The year component of " & theTime.ToString() & " is " & theTime.ToString2("yyyy") & "."
End Sub

' The example produces the following output:
'    The year component of 2/17/2008 9:00:00 AM -07:00 is 2008.
'    The year component of 2/17/2008 9:00:00 AM -07:00 is 8.
'    The year component of 2/17/2008 9:00:00 AM -07:00 is 08.
'    The year component of 2/17/2008 9:00:00 AM -07:00 is 2008.
