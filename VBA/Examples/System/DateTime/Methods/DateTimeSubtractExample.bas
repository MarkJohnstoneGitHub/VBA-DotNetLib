Attribute VB_Name = "DateTimeSubtractExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 13, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.subtract?view=netframework-4.8.1

Option Explicit

'@Description("The following example demonstrates the Subtract method and the subtraction operator.")
Public Sub DateTimeSubtract()
Attribute DateTimeSubtract.VB_Description = "The following example demonstrates the Subtract method and the subtraction operator."
   Dim date1 As DotNetLib.DateTime
   Set date1 = DateTime.CreateFromDateTime(1996, 6, 3, 22, 15, 0)
   Dim date2 As DotNetLib.DateTime
   Set date2 = DateTime.CreateFromDateTime(1996, 12, 6, 13, 2, 0)
   Dim date3 As DotNetLib.DateTime
   Set date3 = DateTime.CreateFromDateTime(1996, 10, 12, 8, 42, 0)
   
   ' diff1 gets 185 days, 14 hours, and 47 minutes.
   Dim diff1 As DotNetLib.TimeSpan
   Set diff1 = date2.Subtract2(date1)
   Debug.Print diff1.ToString()
   
   ' date4 gets 4/9/1996 5:55:00 PM.
   Dim date4 As DotNetLib.DateTime
   Set date4 = date3.Subtract(diff1)
   Debug.Print date4.ToString()
   
   ' diff2 gets 55 days 4 hours and 20 minutes.
   Dim diff2 As DotNetLib.TimeSpan
   Set diff2 = DateTime.Subtraction(date2, date3)
   Debug.Print diff2.ToString()
   
   ' date5 gets 4/9/1996 5:55:00 PM.
   Dim date5 As DotNetLib.DateTime
   Set date5 = DateTime.Subtraction2(date1, diff2)
   Debug.Print date5.ToString()
End Sub

'Output with the culture as English (Australia)
'    185.14:47:00
'    9/04/1996 5:55:00 PM
'    55.04:20:00
'    9/04/1996 5:55:00 PM
