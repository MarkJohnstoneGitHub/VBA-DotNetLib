Attribute VB_Name = "DateTimeOffsetToStringExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tostring?view=netframework-4.8.1#system-datetimeoffset-tostring

Option Explicit

'@Description("The following example illustrates calls to the ToString() method and displays its output on a system whose current culture is en-us.")
Public Sub DateTimeOffsetToString()
Attribute DateTimeOffsetToString.VB_Description = "The following example illustrates calls to the ToString() method and displays its output on a system whose current culture is en-us."
    Dim thisDate As IDateTimeOffset
    Set thisDate = DateTimeOffset.UtcNow
    Debug.Print thisDate.ToString()         ' Displays 3/28/2007 7:13:50 PM +00:00
    
    ' Show output for local time
    Set thisDate = DateTimeOffset.Now
    Debug.Print thisDate.ToString()         ' Displays 3/28/2007 12:13:50 PM -07:00

    ' Show output for arbitrary time offset
    Set thisDate = thisDate.ToOffset(TimeSpan.Create(-5, 0, 0))
    Debug.Print thisDate.ToString()         ' Displays 3/28/2007 2:13:50 PM -05:00
End Sub

' Output :
' 3/28/2007 7:13:50 PM +00:00
' 3/28/2007 12:13:50 PM -07:00
' 3/28/2007 12:13:50 PM -05:00
