Attribute VB_Name = "TimeZoneInfoClearCachedDataEg"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 23, 2023
'@LastModified August 17, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.clearcacheddata?view=netframework-4.8.1#remarks

Option Explicit

'@Description("Cached time zone data includes data on the local time zone and the Coordinated Universal Time (UTC) zone.")
'
' You might call the ClearCachedData method to reduce the memory devoted to the
' application's cache of time zone information or to reflect the fact that the
' local system's time zone has changed.
'
' Storing references to the local and UTC time zones is not recommended.
' After the call to the ClearCachedData method, these object variables will be
' undefined TimeZoneInfo objects that are no longer references to TimeZoneInfo.Local
' or TimeZoneInfo.Utc.
' For example, in the following code, the second call to the
' TimeZoneInfo.ConvertTime(DateTime, TimeZoneInfo, TimeZoneInfo) method throws
' an ArgumentException because the local variable is no longer considered equal
' to TimeZoneInfo.Local.
Public Sub TimeZoneInfoClearCachedData()
Attribute TimeZoneInfoClearCachedData.VB_Description = "Cached time zone data includes data on the local time zone and the Coordinated Universal Time (UTC) zone."
    Dim cst As ITimeZoneInfo
    Set cst = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
    Dim Locale As ITimeZoneInfo
    Set Locale = TimeZoneInfo.Locale
    Debug.Print TimeZoneInfo.ConvertTime3(DateTime.Now, Locale, cst).ToString()
    
    TimeZoneInfo.ClearCachedData
    On Error Resume Next
    Debug.Print TimeZoneInfo.ConvertTime3(DateTime.Now, Locale, cst).ToString()
    If Catch(ArgumentException) Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub
