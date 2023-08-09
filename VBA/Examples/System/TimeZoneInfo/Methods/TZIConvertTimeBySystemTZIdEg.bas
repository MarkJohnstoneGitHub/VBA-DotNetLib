Attribute VB_Name = "TZIConvertTimeBySystemTZIdEg"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 23, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttimebysystemtimezoneid?view=netframework-4.8.1

Option Explicit

'The following example uses the TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime, String, String)
'method to display the time that corresponds to the local system time in eight cities of the world.
Public Sub TimeZoneInfoConvertTimeBySystemTimeZoneId()
    Dim currentTime As IDateTime
    Set currentTime = DateTime.Now
    Debug.Print "Current Times:"
    Debug.Print "Los Angeles: " & _
        TimeZoneInfo.ConvertTimeBySystemTimeZoneId3(currentTime, TimeZoneInfo.Locale.Id, "Pacific Standard Time").ToString()
    Debug.Print "Chicago: " & _
        TimeZoneInfo.ConvertTimeBySystemTimeZoneId3(currentTime, TimeZoneInfo.Locale.Id, "Central Standard Time").ToString()
    Debug.Print "New York: " & _
        TimeZoneInfo.ConvertTimeBySystemTimeZoneId3(currentTime, TimeZoneInfo.Locale.Id, "Eastern Standard Time").ToString()
    Debug.Print "Moscow: " & _
        TimeZoneInfo.ConvertTimeBySystemTimeZoneId3(currentTime, TimeZoneInfo.Locale.Id, "Russian Standard Time").ToString()
    Debug.Print "New Delhi: " & _
        TimeZoneInfo.ConvertTimeBySystemTimeZoneId3(currentTime, TimeZoneInfo.Locale.Id, "India Standard Time").ToString()
    Debug.Print "Beijing: " & _
        TimeZoneInfo.ConvertTimeBySystemTimeZoneId3(currentTime, TimeZoneInfo.Locale.Id, "China Standard Time").ToString()
    Debug.Print "Tokyo: " & _
        TimeZoneInfo.ConvertTimeBySystemTimeZoneId3(currentTime, TimeZoneInfo.Locale.Id, "Tokyo Standard Time").ToString()
End Sub

' Output example:
'
'    Current times:
'    Los Angeles: 31/07/2023 4:21:27 AM
'    Chicago: 31/07/2023 6:21:27 AM
'    New York: 31/07/2023 7:21:27 AM
'    Moscow: 31/07/2023 2:21:27 PM
'    New Delhi: 31/07/2023 4:51:27 PM
'    Beijing: 31/07/2023 7:21:27 PM
'    Tokyo: 31/07/2023 8:21:27 PM
