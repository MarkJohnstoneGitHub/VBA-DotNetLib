Attribute VB_Name = "DateTimeOffsetSubtractExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.subtract?view=netframework-4.8.1#system-datetimeoffset-subtract(system-datetimeoffset)

Option Explicit

''
' The following example illustrates subtraction that uses the
' Subtract(DateTimeOffset) method.
''
Public Sub DateTimeOffsetSubtract()
    Dim firstDate As DotNetLib.DateTimeOffset
    Set firstDate = DateTimeOffset.CreateFromDateTimeParts(2018, 10, 25, 18, 0, 0, TimeSpan.Create(-7, 0, 0))
    Dim secondDate As DotNetLib.DateTimeOffset
    Set secondDate = DateTimeOffset.CreateFromDateTimeParts(2018, 10, 25, 18, 0, 0, TimeSpan.Create(-5, 0, 0))
    Dim thirdDate As DotNetLib.DateTimeOffset
    Set thirdDate = DateTimeOffset.CreateFromDateTimeParts(2018, 9, 28, 9, 0, 0, TimeSpan.Create(-7, 0, 0))
    Dim difference As DotNetLib.TimeSpan
    
    Set difference = firstDate.Subtract(secondDate)
    Debug.Print VBString.Format("({0}) - ({1}): {2} days, {3}:{4:d2}", _
                                firstDate, secondDate, difference.Days, difference.Hours, difference.Minutes)
            
    Set difference = firstDate.Subtract(thirdDate)
    Debug.Print VBString.Format("({0}) - ({1}): {2} days, {3}:{4:d2}", _
                                firstDate, thirdDate, difference.Days, difference.Hours, difference.Minutes)
End Sub

' The example produces the following output:
'    (10/25/2018 6:00:00 PM -07:00) - (10/25/2018 6:00:00 PM -05:00): 0 days, 2:00
'    (10/25/2018 6:00:00 PM -07:00) - (9/28/2018 9:00:00 AM -07:00): 27 days, 9:00

