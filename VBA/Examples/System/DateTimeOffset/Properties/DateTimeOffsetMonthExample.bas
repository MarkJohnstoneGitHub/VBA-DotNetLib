Attribute VB_Name = "DateTimeOffsetMonthExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.month?view=netframework-4.8.1#examples

Option Explicit

''
' The following example displays the month component of a DateTimeOffset value
' in three different ways:
'  By retrieving the value of the Month property.
'  By calling the ToString(String) method with the "M" format specifier.
'  By calling the ToString(String) method with the "MM" format specifier.
''
Public Sub DateTimeOffsetMonth()
    Dim theTime As DotNetLib.DateTimeOffset
    Set theTime = DateTimeOffset.CreateFromDateTimeParts(2008, 9, 7, 11, 25, 0, DateTimeOffset.Now.Offset)
    
    Debug.Print VBString.Format("The month component of {0} is {1}.", _
                                theTime, theTime.Month)

    Debug.Print VBString.Format("The month component of {0} is{1}.", _
                                theTime, theTime.ToString2(" M"))

    Debug.Print VBString.Format("The month component of {0} is {1}.", _
                                theTime, theTime.ToString2("MM"))
End Sub

' The example produces the following output:
'    The month component of 9/7/2008 11:25:00 AM -08:00 is 9.
'    The month component of 9/7/2008 11:25:00 AM -08:00 is 9.
'    The month component of 9/7/2008 11:25:00 AM -08:00 is 09.

