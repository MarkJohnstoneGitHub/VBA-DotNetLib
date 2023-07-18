Attribute VB_Name = "DateTimeOffsetMonthExample"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified July 19, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.month?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example displays the month component of a DateTimeOffset value in three different ways:")
'  By retrieving the value of the Month property.
'  By calling the ToString(String) method with the "M" format specifier.
'  By calling the ToString(String) method with the "MM" format specifier.
Public Sub DateTimeOffsetMonth()
Attribute DateTimeOffsetMonth.VB_Description = "The following example displays the month component of a DateTimeOffset value in three different ways:"
   Dim theTime As DateTimeOffset
   Set theTime = DateTimeOffset.CreateFromDateTimeParts(2008, 9, 7, 11, 25, 0, DateTimeOffset.Now.Offset)
   Debug.Print "The month component of " & theTime.ToString() & " is " & theTime.Month & "."
   Debug.Print "The month component of " & theTime.ToString() & " is" & theTime.ToString2(" M") & "."
   Debug.Print "The month component of " & theTime.ToString() & " is " & theTime.ToString2("MM") & "."

' The example produces the following output:
'    The month component of 9/7/2008 11:25:00 AM -08:00 is 9.
'    The month component of 9/7/2008 11:25:00 AM -08:00 is 9.
'    The month component of 9/7/2008 11:25:00 AM -08:00 is 09.
End Sub
