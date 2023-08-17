Attribute VB_Name = "DateTimeOffsetUtcDateTimeEg"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.utcdatetime?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example shows how to use of the UtcDateTime property to display a DateTimeOffset value and its corresponding UTC time.")
Public Sub DateTimeOffsetUtcDateTime()
Attribute DateTimeOffsetUtcDateTime.VB_Description = "The following example shows how to use of the UtcDateTime property to display a DateTimeOffset value and its corresponding UTC time."
    Dim offsetTime As IDateTimeOffset
    Set offsetTime = DateTimeOffset.CreateFromDateTimeParts(2007, 11, 25, 11, 14, 0, TimeSpan.Create(3, 0, 0))
    Debug.Print offsetTime.ToString() & " is equivalent to " & _
                offsetTime.UtcDateTime.ToString() & " " & _
                DateTimeKindHelper.ToString(offsetTime.UtcDateTime.Kind)
End Sub

' The example displays the following output:
'       11/25/2007 11:14:00 AM +03:00 is equivalent to 11/25/2007 8:14:00 AM Utc
