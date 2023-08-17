Attribute VB_Name = "DateTimeOffsetAddDaysExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 20, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.adddays?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the AddDays method to list the dates that fall on Monday, the start of the work week, in March 2008.")
Public Sub DateTimeOffsetAddDays()
Attribute DateTimeOffsetAddDays.VB_Description = "The following example uses the AddDays method to list the dates that fall on Monday, the start of the work week, in March 2008."
   Dim workDay As IDateTimeOffset
   Set workDay = DateTimeOffset.CreateFromDateTimeParts(2008, 3, 1, 9, 0, 0, DateTimeOffset.Now.Offset)
   Dim pvtMonth As Long
   pvtMonth = workDay.Month
   
   ' Start with the first Monday of the month
   If (workDay.DayOfWeek <> DayOfWeek.DayOfWeek_Monday) Then
      If (workDay.DayOfWeek = DayOfWeek.DayOfWeek_Sunday) Then
         Set workDay = workDay.AddDays(1)
      Else
         Set workDay = workDay.AddDays(8 - workDay.DayOfWeek)
      End If
   End If
   Debug.Print "Beginning of Work Week In " & workDay.ToString2("MMMM") & " " & workDay.ToString2("yyyy") & ":"
   
   ' Add one week to the current date
   Do
      Debug.Print "   " & workDay.ToString2("dddd") & ", " & workDay.ToString2("MMMM") & workDay.ToString2(" d")
      Set workDay = workDay.AddDays(7)
   Loop While (workDay.Month = pvtMonth)
End Sub

' The example produces the following output:
'    Beginning of Work Week In March 2008:
'       Monday, March 3
'       Monday, March 10
'       Monday, March 17
'       Monday, March 24
'       Monday, March 31
