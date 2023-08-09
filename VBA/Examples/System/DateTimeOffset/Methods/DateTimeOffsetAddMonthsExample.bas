Attribute VB_Name = "DateTimeOffsetAddMonthsExample"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 20, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.addmonths?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the AddMonths method to display the start date of each quarter of the year 2007.")
Public Sub DateTimeOffsetAddMonths()
Attribute DateTimeOffsetAddMonths.VB_Description = "The following example uses the AddMonths method to display the start date of each quarter of the year 2007."
   Dim quarterDate As IDateTimeOffset
   Set quarterDate = DateTimeOffset.CreateFromDateTimeParts(2007, 1, 1, 0, 0, 0, DateTimeOffset.Now.Offset)
   
   Dim ctr As Long
   For ctr = 1 To 4
      Debug.Print "Quarter " & ctr & ": " & quarterDate.ToString2("MMMM d")
      Set quarterDate = quarterDate.AddMonths(3)
   Next
End Sub

' This example produces the following output:
'       Quarter 1: January 1
'       Quarter 2: April 1
'       Quarter 3: July 1
'       Quarter 4: October 1
