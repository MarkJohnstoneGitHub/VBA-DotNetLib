Attribute VB_Name = "DateTimeToLongDateStringExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified July 14, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tolongdatestring?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the ToLongDateString method.")
Public Sub DateTimeToLongDateString()
Attribute DateTimeToLongDateString.VB_Description = "The following example demonstrates the ToLongDateString method."
   ' Initialize a DateTime object.
   Debug.Print "Initialize the DateTime object to May 16, 2001 3:02:15 AM." & VBA.vbNewLine
   Dim dateAndTime As DateTime
   Set dateAndTime = DateTime.CreateFromDateTime(2001, 5, 16, 3, 2, 15)
   
   Debug.Print "Long date string:  " & dateAndTime.ToLongDateString()
   Debug.Print "Long time string:  " & dateAndTime.ToLongTimeString()
   Debug.Print "Short date string: " & dateAndTime.ToShortDateString()
   Debug.Print "Short time string: " & dateAndTime.ToShortTimeString()
   
End Sub
