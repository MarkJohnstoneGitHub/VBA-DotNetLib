Attribute VB_Name = "DateTimeToLongDateStringExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tolongdatestring?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the ToLongDateString method.")
Public Sub DateTimeToLongDateString()
Attribute DateTimeToLongDateString.VB_Description = "The following example demonstrates the ToLongDateString method."
   ' Initialize a DateTime object.
   Debug.Print "Initialize the DateTime object to May 16, 2001 3:02:15 AM." & VBA.vbNewLine
   Dim dateAndTime As DotNetLib.DateTime
   Set dateAndTime = DateTime.CreateFromDateTime(2001, 5, 16, 3, 2, 15)
   
   Debug.Print "Long date string:  " & dateAndTime.ToLongDateString()
   Debug.Print "Long time string:  " & dateAndTime.ToLongTimeString()
   Debug.Print "Short date string: " & dateAndTime.ToShortDateString()
   Debug.Print "Short time string: " & dateAndTime.ToShortTimeString()
End Sub

' The example displays output similar to the following:
'        Current culture: "en-US"
'
'        Long date pattern: "dddd, MMMM d, yyyy"
'        Long date string:  "Wednesday, May 16, 2001"
'
'        Long time pattern: "h:mm:ss tt"
'        Long time string:  "3:02:15 AM"
'
'        Short date pattern: "M/d/yyyy"
'        Short date string:  "5/16/2001"
'
'        Short time pattern: "h:mm tt"
'        Short time string:  "3:02 AM"
