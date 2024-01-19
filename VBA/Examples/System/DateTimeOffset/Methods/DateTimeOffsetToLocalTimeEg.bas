Attribute VB_Name = "DateTimeOffsetToLocalTimeEg"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tolocaltime?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the ToLocalTime method to convert a DateTimeOffset
' value to local time in the Pacific Standard Time zone.
''
Public Sub DateTimeOffsetToLocalTime()
    Dim originalTime As DotNetLib.DateTimeOffset
    Dim localTime As DotNetLib.DateTimeOffset
    
    Set originalTime = DateTimeOffset.CreateFromDateTimeParts(2007, 3, 11, 3, 0, 0, TimeSpan.Create(-6, 0, 0))
    Set localTime = originalTime.ToLocalTime()
    Debug.Print VBString.Format("Converted {0} to {1}.", originalTime.ToString(), _
                                localTime.ToString())
    
    Set originalTime = DateTimeOffset.CreateFromDateTimeParts(2007, 3, 11, 4, 0, 0, TimeSpan.Create(-6, 0, 0))
    Set localTime = originalTime.ToLocalTime()
    Debug.Print VBString.Format("Converted {0} to {1}.", originalTime.ToString(), _
                                localTime.ToString())
                                           
    ' Define a summer UTC time
    Set originalTime = DateTimeOffset.CreateFromDateTimeParts(2007, 6, 15, 8, 0, 0, TimeSpan.Zero)
    Set localTime = originalTime.ToLocalTime()
    Debug.Print VBString.Format("Converted {0} to {1}.", originalTime.ToString(), _
                                localTime.ToString())
                
    ' Define a winter time
    Set originalTime = DateTimeOffset.CreateFromDateTimeParts(2007, 11, 30, 14, 0, 0, TimeSpan.Create(3, 0, 0))
    Set localTime = originalTime.ToLocalTime()
    Debug.Print VBString.Format("Converted {0} to {1}.", originalTime.ToString(), _
                                localTime.ToString())
End Sub

' The example produces the following output:
'    Converted 3/11/2007 3:00:00 AM -06:00 to 3/11/2007 1:00:00 AM -08:00.
'    Converted 3/11/2007 4:00:00 AM -06:00 to 3/11/2007 3:00:00 AM -07:00.
'    Converted 6/15/2007 8:00:00 AM +00:00 to 6/15/2007 1:00:00 AM -07:00.
'    Converted 11/30/2007 2:00:00 PM +03:00 to 11/30/2007 3:00:00 AM -08:00.

