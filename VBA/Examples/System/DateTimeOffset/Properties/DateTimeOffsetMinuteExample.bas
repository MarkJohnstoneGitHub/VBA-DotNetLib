Attribute VB_Name = "DateTimeOffsetMinuteExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.minute?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example displays the minute component of a DateTimeOffset object in three different ways:")
'  By retrieving the value of the Minute property.
'  By calling the ToString(String) method with the "m" format specifier.
'  By calling the ToString(String) method with the "mm" format specifier.
Public Sub DateTimeOffsetMinute()
Attribute DateTimeOffsetMinute.VB_Description = "The following example displays the minute component of a DateTimeOffset object in three different ways:"
    Dim theTime As IDateTimeOffset
    Set theTime = DateTimeOffset.CreateFromDateTimeParts(2008, 5, 1, 10, 3, 0, DateTimeOffset.Now.Offset)
    Debug.Print "The minute component of " & theTime.ToString() & " is " & theTime.Minute & "."
    Debug.Print "The minute component of " & theTime.ToString() & " is" & theTime.ToString2(" m") & "."
    Debug.Print "The minute component of " & theTime.ToString() & " is " & theTime.ToString2("mm") & "."
End Sub

' The example produces the following output:
'    The minute component of 5/1/2008 10:03:00 AM -08:00 is 3.
'    The minute component of 5/1/2008 10:03:00 AM -08:00 is 3.
'    The minute component of 5/1/2008 10:03:00 AM -08:00 is 03.
