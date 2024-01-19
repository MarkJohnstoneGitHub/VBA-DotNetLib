Attribute VB_Name = "DateTimeOffsetSecondExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.second?view=netframework-4.8.1#examples

Option Explicit

''
' The following example displays the second component of a DateTimeOffset object
' in three different ways:
'  By retrieving the value of the Second property.
'  By calling the ToString(String) method with the "s" format specifier.
'  By calling the ToString(String) method with the "ss" format specifier.
''
Public Sub DateTimeOffsetSecond()
    Dim theTime As DotNetLib.DateTimeOffset
    Set theTime = DateTimeOffset.CreateFromDateTimeParts(2008, 6, 12, 21, 16, 32, DateTimeOffset.Now.Offset)
    Debug.Print VBString.Format("The second component of {0} is {1}.", _
                                theTime, theTime.SECOND)

    Debug.Print VBString.Format("The second component of {0} is{1}.", _
                                theTime, theTime.ToString2(" s"))

    Debug.Print VBString.Format("The second component of {0} is {1}.", _
                                theTime, theTime.ToString2("ss"))
End Sub

' The example produces the following output:
'    The second component of 6/12/2008 9:16:32 PM -07:00 is 32.
'    The second component of 6/12/2008 9:16:32 PM -07:00 is 32.
'    The second component of 6/12/2008 9:16:32 PM -07:00 is 32.

