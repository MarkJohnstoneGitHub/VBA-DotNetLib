Attribute VB_Name = "DateTimeOffsetOffsetExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.offset?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the Offset property to display the local time's
' difference from Coordinated Universal Time (UTC).
''
Public Sub DateTimeOffsetOffset()
    Dim localTime As DotNetLib.DateTimeOffset
    Set localTime = DateTimeOffset.Now
    Debug.Print VBString.Format("The local time zone is {0} hours and {1} minutes {2} than UTC.", _
                                Abs(localTime.Offset.Hours), _
                                localTime.Offset.Minutes, _
                                IIf(localTime.Offset.Hours < 0, "earlier", "later"))
End Sub

' The example displays output similar to the following for a system in the
' U.S. Pacific Standard Time zone:
'       The local time zone is 8 hours and 0 minutes earlier than UTC.

