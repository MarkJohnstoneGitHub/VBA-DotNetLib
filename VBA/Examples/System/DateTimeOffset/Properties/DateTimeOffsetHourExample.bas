Attribute VB_Name = "DateTimeOffsetHourExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.hour?view=netframework-4.8.1#examples

Option Explicit

''
' The following example displays the hour component of a DateTimeOffset object
' in three different ways:
'  By retrieving the value of the Hour property.
'  By calling the ToString(String) method with the "H" format specifier.
'  By calling the ToString(String) method with the "HH" format specifier.
''
Public Sub DateTimeOffsetHour()
    Dim theTime As DotNetLib.DateTimeOffset
    Set theTime = DateTimeOffset.CreateFromDateTimeParts(2008, 3, 1, 14, 15, 0, DateTimeOffset.Now.Offset)
    
    Debug.Print VBString.Format("The hour component of {0} is {1}.", _
                                theTime, theTime.Hour)

    Debug.Print VBString.Format("The hour component of {0} is{1}.", _
                                theTime, theTime.ToString2(" H"))

    Debug.Print VBString.Format("The hour component of {0} is {1}.", _
                                theTime, theTime.ToString2("HH"))
End Sub
   
' The example produces the following output:
'    The hour component of 3/1/2008 2:15:00 PM -08:00 is 14.
'    The hour component of 3/1/2008 2:15:00 PM -08:00 is 14.
'    The hour component of 3/1/2008 2:15:00 PM -08:00 is 14.

