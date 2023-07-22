Attribute VB_Name = "TimeZoneInfoClearCachedDataEg"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 23, 2023
'@LastModified July 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.clearcacheddata?view=netframework-4.8.1#remarks

Option Explicit

'@Description("Cached time zone data includes data on the local time zone and the Coordinated Universal Time (UTC) zone.")
Public Sub TimeZoneInfoClearCachedData()
Attribute TimeZoneInfoClearCachedData.VB_Description = "Cached time zone data includes data on the local time zone and the Coordinated Universal Time (UTC) zone."
    Dim cst As TimeZoneInfo
    Set cst = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
    Dim locale As TimeZoneInfo
    Set locale = TimeZoneInfo.locale
    Debug.Print TimeZoneInfo.ConvertTime3(DateTime.Now, locale, cst).ToString()
    
    TimeZoneInfo.ClearCachedData
    On Error Resume Next
    Debug.Print TimeZoneInfo.ConvertTime3(DateTime.Now, locale, cst).ToString()
    If Catch(ArgumentException) Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
   
End Sub


'TimeZoneInfo cst = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");
'TimeZoneInfo local = TimeZoneInfo.Local;
'Console.WriteLine(TimeZoneInfo.ConvertTime(DateTime.Now, local, cst));
'
'TimeZoneInfo.ClearCachedData();
'Try
'{
'   Console.WriteLine(TimeZoneInfo.ConvertTime(DateTime.Now, local, cst));
'}
'catch (ArgumentException e)
'{
'   Console.WriteLine(e.GetType().Name + "\n   " + e.Message);
'}
