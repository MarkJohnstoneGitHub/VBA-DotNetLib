Attribute VB_Name = "DateTimeAddMonthsExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified August 4, 2023

'Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.addmonths?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example adds between zero and fifteen months to the last day of December, 2015. In this case, the AddMonths method returns the date of the last day of each month, and successfully handles leap years.")
Public Sub DateTimeAddMonths()
Attribute DateTimeAddMonths.VB_Description = "The following example adds between zero and fifteen months to the last day of December, 2015. In this case, the AddMonths method returns the date of the last day of each month, and successfully handles leap years."
   Dim dat As IDateTime
   Set dat = DateTime.CreateFromDate(2015, 12, 31)
   Dim ctr As Long
   For ctr = 0 To 15
      Debug.Print dat.AddMonths(ctr).ToString2("d")
   Next
End Sub

' The example displays the following output:
'       12/31/2015
'       1/31/2016
'       2/29/2016
'       3/31/2016
'       4/30/2016
'       5/31/2016
'       6/30/2016
'       7/31/2016
'       8/31/2016
'       9/30/2016
'       10/31/2016
'       11/30/2016
'       12/31/2016
'       1/31/2017
'       2/28/2017
'       3/31/2017
