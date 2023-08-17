Attribute VB_Name = "DTOToUnixTimeSecondsExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tounixtimeseconds?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example calls the ToUnixTimeSeconds method to return the Unix time of values that are equal to, shortly before, and shortly after 1970-01-01T00:00:00Z.")
Public Sub DateTimeOffsetToUnixTimeSeconds()
Attribute DateTimeOffsetToUnixTimeSeconds.VB_Description = "The following example calls the ToUnixTimeSeconds method to return the Unix time of values that are equal to, shortly before, and shortly after 1970-01-01T00:00:00Z."
    Dim dto As IDateTimeOffset
    Set dto = DateTimeOffset.CreateFromDateTimeParts(1970, 1, 1, 0, 0, 0, TimeSpan.Zero)
    Debug.Print dto.ToString() & " --> Unix Seconds: " & dto.ToUnixTimeSeconds()

    Set dto = DateTimeOffset.CreateFromDateTimeParts(1969, 12, 31, 23, 59, 0, TimeSpan.Zero)
    Debug.Print dto.ToString() & " --> Unix Seconds: " & dto.ToUnixTimeSeconds()
    
    Set dto = DateTimeOffset.CreateFromDateTimeParts(1970, 1, 1, 0, 1, 0, TimeSpan.Zero)
    Debug.Print dto.ToString() & " --> Unix Seconds: " & dto.ToUnixTimeSeconds()
End Sub

' The example displays the following output:
'    1/1/1970 12:00:00 AM +00:00 --> Unix Seconds: 0
'    12/31/1969 11:59:00 PM +00:00 --> Unix Seconds: -60
'    1/1/1970 12:01:00 AM +00:00 --> Unix Seconds: 60
