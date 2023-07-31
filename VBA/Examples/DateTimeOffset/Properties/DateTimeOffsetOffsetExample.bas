Attribute VB_Name = "DateTimeOffsetOffsetExample"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.offset?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the Offset property to display the local time's difference from Coordinated Universal Time (UTC).")
Public Sub DateTimeOffsetOffset()
Attribute DateTimeOffsetOffset.VB_Description = "The following example uses the Offset property to display the local time's difference from Coordinated Universal Time (UTC)."
    Dim localTime As IDateTimeOffset
    Set localTime = DateTimeOffset.Now
    Debug.Print "The local time zone is " & _
                Abs(localTime.Offset.Hours) & " hours and " & _
                localTime.Offset.Minutes & " minutes " & _
                IIf(localTime.Offset.Hours < 0, "earlier", "later") & " than UTC."
End Sub

' The example displays output similar to the following for a system in the
' U.S. Pacific Standard Time zone:
'       The local time zone is 8 hours and 0 minutes earlier than UTC.
