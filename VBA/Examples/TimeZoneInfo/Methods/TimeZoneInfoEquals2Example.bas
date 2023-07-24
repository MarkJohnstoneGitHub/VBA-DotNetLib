Attribute VB_Name = "TimeZoneInfoEquals2Example"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 24, 2023
'@LastModified July 24, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.equals?view=netframework-4.8.1#system-timezoneinfo-equals(system-object)

Option Explicit

'@Description("The following example uses the Equals(Object) method to determine whether the local time zone is Pacific Time or Eastern Time.")
Public Sub TimeZoneInfoEquals2()
Attribute TimeZoneInfoEquals2.VB_Description = "The following example uses the Equals(Object) method to determine whether the local time zone is Pacific Time or Eastern Time."
    Dim thisTimeZone As TimeZoneInfo
    Dim obj1 As Object
    Dim obj2 As Object
    
    Set thisTimeZone = TimeZoneInfo.Locale
    Set obj1 = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time")
    Set obj2 = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")
    Debug.Print thisTimeZone.Equals2(obj1)
    Debug.Print thisTimeZone.Equals2(obj2)
End Sub

' The example displays the following output:
'      True
'      False
