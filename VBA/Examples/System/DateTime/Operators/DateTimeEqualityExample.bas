Attribute VB_Name = "DateTimeEqualityExample"
'@Folder "Examples.System.DateTime.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.op_equality?view=netframework-4.8.1#examples

Option Explicit

Public Sub DateTimeEquality()
   Dim april19 As IDateTime
   Set april19 = DateTime.CreateFromDate(2001, 4, 19)
   Dim otherDate As IDateTime
   Set otherDate = DateTime.CreateFromDate(1991, 6, 5)
   
   ' areEqual gets false.
   Dim areEqual As Boolean
   areEqual = DateTime.Equality(april19, otherDate)
   Debug.Print april19.ToString & " = " & otherDate.ToString & " is " & areEqual

   Set otherDate = DateTime.CreateFromDate(2001, 4, 19)
   ' areEqual gets true.
   areEqual = DateTime.Equality(april19, otherDate)
   Debug.Print april19.ToString & " = " & otherDate.ToString & " is " & areEqual
End Sub

' Output:
'    19/04/2001 12:00:00 AM = 5/06/1991 12:00:00 AM is False
'    19/04/2001 12:00:00 AM = 19/04/2001 12:00:00 AM is True
