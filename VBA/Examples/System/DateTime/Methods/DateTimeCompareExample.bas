Attribute VB_Name = "DateTimeCompareExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.compare?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the Compare method.")
Public Sub DateTimeCompare()
Attribute DateTimeCompare.VB_Description = "The following example demonstrates the Compare method."
   Dim date1 As IDateTime
   Set date1 = DateTime.CreateFromDateTime(2009, 8, 1, 0, 0, 0)
   Dim date2 As IDateTime
   Set date2 = DateTime.CreateFromDateTime(2009, 8, 1, 12, 0, 0)
   
   Dim Result As Long
   Result = DateTime.Compare(date1, date2)
   Dim relationship As String

   If (Result < 0) Then
      relationship = "is earlier than"
   ElseIf (Result = 0) Then
      relationship = "is the same time as"
   Else
      relationship = "is later than"
   End If
   Debug.Print date1.ToString & " " & relationship & " " & date2.ToString
End Sub

' The example displays the following output for en-us culture:
'    8/1/2009 12:00:00 AM is earlier than 8/1/2009 12:00:00 PM
