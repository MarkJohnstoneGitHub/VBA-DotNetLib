Attribute VB_Name = "DateTimeAddDaysExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.adddays?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the AddDays method to determine the day of the week 36 days after the current date.")
Public Sub DateTimeAddDays()
Attribute DateTimeAddDays.VB_Description = "The following example uses the AddDays method to determine the day of the week 36 days after the current date."
    Dim pvtToday As DotNetLib.DateTime
    Set pvtToday = DateTime.Now
    Dim answer As DotNetLib.DateTime
    Set answer = pvtToday.AddDays(36)
    Debug.Print VBString.Format("Today: {0:dddd}", pvtToday)
    Debug.Print VBString.Format("36 days from today: {0:dddd}", answer)
End Sub

' The example displays output like the following:
'       Today: Wednesday
'       36 days from today: Thursday
