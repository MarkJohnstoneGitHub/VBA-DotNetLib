Attribute VB_Name = "DateTimeOffsetAddMonthsExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 20, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.addmonths?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the AddMonths method to display the start date of
' each quarter of the year 2007.
''
Public Sub DateTimeOffsetAddMonths()
    Dim quarterDate As DotNetLib.DateTimeOffset
    Set quarterDate = DateTimeOffset.CreateFromDateTimeParts(2007, 1, 1, 0, 0, 0, DateTimeOffset.Now.Offset)
    Dim ctr As Long
    For ctr = 1 To 4
        Debug.Print VBString.Format("Quarter {0}: {1:MMMM d}", ctr, quarterDate)
        Set quarterDate = quarterDate.AddMonths(3)
    Next
End Sub

' This example produces the following output:
'       Quarter 1: January 1
'       Quarter 2: April 1
'       Quarter 3: July 1
'       Quarter 4: October 1
