Attribute VB_Name = "DateTimeAddDaysExample"
'Rubberduck annotations
'@Folder "VBADotNetLib.Examples.DateTime.Methods"

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 09, 2023

'@DotNetReference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.adddays?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the AddDays method to determine the day of the week 36 days after the current date.")
Public Sub DateTimeAddDays()
   Dim pvtToday As DateTime
   Set pvtToday = DateTime.Now
   Dim answer As DateTime
   Set answer = pvtToday.AddDays(36)
   Debug.Print "Today: " & pvtToday.ToString2("dddd")
   Debug.Print "36 days from today: " & answer.ToString2("dddd")
   
' The example displays output like the following:
'       Today: Wednesday
'       36 days from today: Thursday

End Sub
