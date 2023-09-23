Attribute VB_Name = "DateTimeAddDaysExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified September 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.adddays?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the AddDays method to determine the day of the week 36 days after the current date.")
Public Sub DateTimeAddDays()
Attribute DateTimeAddDays.VB_Description = "The following example uses the AddDays method to determine the day of the week 36 days after the current date."
    Dim pvtToday As IDateTime
    Set pvtToday = DateTime.Now
    Dim answer As IDateTime
    Set answer = pvtToday.AddDays(36)
    Debug.Print Strings.Format("Today: {0:dddd}", pvtToday)
    Debug.Print Strings.Format("36 days from today: {0:dddd}", answer)
End Sub

' The example displays output like the following:
'       Today: Wednesday
'       36 days from today: Thursday
