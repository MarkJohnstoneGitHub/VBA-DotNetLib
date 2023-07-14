Attribute VB_Name = "DateTimeEqualityExample"
'@Folder("VBADotNetLib.Examples.DateTime.Operators")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified July 14, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.op_equality?view=netframework-4.8.1#examples

Option Explicit

Public Sub DateTimeEquality()
   Dim april19 As DateTime
   Set april19 = DateTime.CreateFromDate(2001, 4, 19)
   Dim otherDate As DateTime
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

