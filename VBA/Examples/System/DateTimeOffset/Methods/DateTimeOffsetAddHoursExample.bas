Attribute VB_Name = "DateTimeOffsetAddHoursExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 20, 2023
'@LastModified January 9, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.addhours?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the AddHours method to list the start times of work
' shifts for a particular week at an office that has two eight-hour shifts per day.
''
Public Sub DateTimeOffsetAddHours()
    Const SHIFT_LENGTH As Long = 8
    
    Dim startTime As DotNetLib.DateTimeOffset
    Set startTime = DateTimeOffset.CreateFromDateTimeParts(2007, 8, 6, 0, 0, 0, DateTimeOffset.Now.Offset)
    Dim startOfShift As DotNetLib.DateTimeOffset
    Set startOfShift = startTime.AddHours(SHIFT_LENGTH)
    Debug.Print VBString.Format("Shifts for the week of {0:D}", startOfShift)
    Do
       ' Exclude third shift
        If (startOfShift.Hour > 6) Then
            Debug.Print VBString.Format("   {0:d} at {0:T}", startOfShift)
        End If
        Set startOfShift = startOfShift.AddHours(SHIFT_LENGTH)
    Loop While (startOfShift.DayOfWeek <> DayOfWeek.DayOfWeek_Saturday And startOfShift.DayOfWeek <> DayOfWeek.DayOfWeek_Sunday)
End Sub

' The example produces the following output:
'
'    Shifts for the week of Monday, August 06, 2007
'       8/6/2007 at 8:00:00 AM
'       8/6/2007 at 4:00:00 PM
'       8/7/2007 at 8:00:00 AM
'       8/7/2007 at 4:00:00 PM
'       8/8/2007 at 8:00:00 AM
'       8/8/2007 at 4:00:00 PM
'       8/9/2007 at 8:00:00 AM
'       8/9/2007 at 4:00:00 PM
'       8/10/2007 at 8:00:00 AM
'       8/10/2007 at 4:00:00 PM
