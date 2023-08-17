Attribute VB_Name = "DateTimeOffsetDayExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.day?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example displays the day component of a DateTimeOffset object in three different ways:")
'  By retrieving the value of the Day property.
'  By calling the ToString(String) method with the "d" format specifier.
'  By calling the ToString(String) method with the "dd" format specifier.
Public Sub DateTimeOffsetDay()
Attribute DateTimeOffsetDay.VB_Description = "The following example displays the day component of a DateTimeOffset object in three different ways:"
    Dim theTime As IDateTimeOffset
    Set theTime = DateTimeOffset.CreateFromDateTimeParts(2007, 5, 1, 16, 35, 0, DateTimeOffset.Now.Offset)
    Debug.Print "The day component of " & theTime.ToString() & " is " & theTime.Day & "."
    Debug.Print "The day component of " & theTime.ToString() & " is" & theTime.ToString2(" d") & "."
    Debug.Print "The day component of " & theTime.ToString() & " is " & theTime.ToString2("dd") & "."
End Sub

' The example produces the following output:
'    The day component of 5/1/2007 4:35:00 PM -08:00 is 1.
'    The day component of 5/1/2007 4:35:00 PM -08:00 is 1.
'    The day component of 5/1/2007 4:35:00 PM -08:00 is 01.
