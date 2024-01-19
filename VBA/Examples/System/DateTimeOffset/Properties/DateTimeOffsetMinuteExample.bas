Attribute VB_Name = "DateTimeOffsetMinuteExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.minute?view=netframework-4.8.1#examples

Option Explicit

''
' The following example displays the minute component of a DateTimeOffset object
' in three different ways:
'  By retrieving the value of the Minute property.
'  By calling the ToString(String) method with the "m" format specifier.
'  By calling the ToString(String) method with the "mm" format specifier.
''
Public Sub DateTimeOffsetMinute()
    Dim theTime As DotNetLib.DateTimeOffset
    Set theTime = DateTimeOffset.CreateFromDateTimeParts(2008, 5, 1, 10, 3, 0, DateTimeOffset.Now.Offset)
    
    Debug.Print VBString.Format("The minute component of {0} is {1}.", _
                                theTime, theTime.Minute)

    Debug.Print VBString.Format("The minute component of {0} is{1}.", _
                                theTime, theTime.ToString2(" m"))

    Debug.Print VBString.Format("The minute component of {0} is {1}.", _
                                theTime, theTime.ToString2("mm"))
End Sub

' The example produces the following output:
'    The minute component of 5/1/2008 10:03:00 AM -08:00 is 3.
'    The minute component of 5/1/2008 10:03:00 AM -08:00 is 3.
'    The minute component of 5/1/2008 10:03:00 AM -08:00 is 03.

