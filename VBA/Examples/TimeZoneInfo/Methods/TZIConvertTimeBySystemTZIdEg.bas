Attribute VB_Name = "TZIConvertTimeBySystemTZIdEg"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 23, 2023
'@LastModified July 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttimebysystemtimezoneid?view=netframework-4.8.1

Option Explicit

'The following example uses the TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime, String, String)
'method to display the time that corresponds to the local system time in eight cities of the world.
Public Sub TimeZoneInfoConvertTimeBySystemTimeZoneId()
    Dim currentTime As DateTime
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

