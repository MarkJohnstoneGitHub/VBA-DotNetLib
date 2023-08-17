Attribute VB_Name = "DateTimeDaysInMonthExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.daysinmonth?view=netframework-4.8.1#examples

Option Explicit

' The following example demonstrates how to use the DaysInMonth method to determine
' the number of days in July 2001, February 1998 (a non-leap year), and February 1996 (a leap year).
Public Sub DateTimeDaysInMonth()
   Const July As Long = 7
   Const Feb As Long = 2
   
   Dim daysInJuly As Long
   daysInJuly = DateTime.DaysInMonth(2001, July)
   Debug.Print daysInJuly
   
   ' daysInFeb gets 28 because the year 1998 was not a leap year.
   Dim daysInFeb As Long
   daysInFeb = DateTime.DaysInMonth(1998, Feb)
   Debug.Print daysInFeb

   ' daysInFebLeap gets 29 because the year 1996 was a leap year.
   Dim daysInFebLeap As Long
   daysInFebLeap = DateTime.DaysInMonth(1996, Feb)
   Debug.Print daysInFebLeap
End Sub

' The example displays the following output:
'       31
'       28
'       29
