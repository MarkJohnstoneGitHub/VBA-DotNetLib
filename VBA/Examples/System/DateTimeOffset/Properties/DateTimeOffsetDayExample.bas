Attribute VB_Name = "DateTimeOffsetDayExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.day?view=netframework-4.8.1#examples

Option Explicit

''
' The following example displays the day component of a DateTimeOffset object
' in three different ways:
'  By retrieving the value of the Day property.
'  By calling the ToString(String) method with the "d" format specifier.
'  By calling the ToString(String) method with the "dd" format specifier.
''
Public Sub DateTimeOffsetDay()
    Dim theTime As DotNetLib.DateTimeOffset
    Set theTime = DateTimeOffset.CreateFromDateTimeParts(2007, 5, 1, 16, 35, 0, DateTimeOffset.Now.Offset)
    
    Debug.Print VBString.Format("The day component of {0} is {1}.", _
                                theTime, theTime.Day)

    Debug.Print VBString.Format("The day component of {0} is{1}.", _
                                theTime, theTime.ToString2(" d"))

    Debug.Print VBString.Format("The day component of {0} is {1}.", _
                                theTime, theTime.ToString2("dd"))
End Sub

' The example produces the following output:
'    The day component of 5/1/2007 4:35:00 PM -08:00 is 1.
'    The day component of 5/1/2007 4:35:00 PM -08:00 is 1.
'    The day component of 5/1/2007 4:35:00 PM -08:00 is 01.

