Attribute VB_Name = "DateTimeIsLeapYearExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 11, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.isleapyear?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the IsLeapYear method to determine which years between 1994 and 2014 are leap years.")
'The example also illustrates the result when the AddYears method is used to add a year to a leap day.
Public Sub DateTimeIsLeapYear()
Attribute DateTimeIsLeapYear.VB_Description = "The following example uses the IsLeapYear method to determine which years between 1994 and 2014 are leap years."
   Dim pvtYear As Long
   For pvtYear = 1994 To 2014
      If (DateTime.IsLeapYear(pvtYear)) Then
         Debug.Print pvtYear & " is a leap year."
         Dim leapDay As IDateTime
         Set leapDay = DateTime.CreateFromDate(pvtYear, 2, 29)
         Dim nextYear As IDateTime
         Set nextYear = leapDay.AddYears(1)
         Debug.Print "   One year from " & leapDay.ToString2("d") & " is " & nextYear.ToString2("d")
      End If
   Next
End Sub

' The example produces the following output:
'       1996 is a leap year.
'          One year from 2/29/1996 is 2/28/1997.
'       2000 is a leap year.
'          One year from 2/29/2000 is 2/28/2001.
'       2004 is a leap year.
'          One year from 2/29/2004 is 2/28/2005.
'       2008 is a leap year.
'          One year from 2/29/2008 is 2/28/2009.
'       2012 is a leap year.
'          One year from 2/29/2012 is 2/28/2013.
