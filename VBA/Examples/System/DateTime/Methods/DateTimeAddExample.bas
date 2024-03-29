Attribute VB_Name = "DateTimeAddExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.add?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the Add method. It calculates the day of the week that is 36 days (864 hours) from this moment.")
Public Sub DateTimeAdd()
Attribute DateTimeAdd.VB_Description = "The following example demonstrates the Add method. It calculates the day of the week that is 36 days (864 hours) from this moment."
    Dim pvtToday As DotNetLib.DateTime
    Set pvtToday = DateTime.Now
    
    Dim pvtDuration As DotNetLib.TimeSpan
    Set pvtDuration = TimeSpan.Create2(36, 0, 0, 0)
    
    Dim answer As IDateTime
    Set answer = pvtToday.Add(pvtDuration)
    Debug.Print answer.ToString()
End Sub
