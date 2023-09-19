Attribute VB_Name = "DateTimeOffsetSubtractExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.subtract?view=netframework-4.8.1#system-datetimeoffset-subtract(system-datetimeoffset)

Option Explicit

'@Description("The following example illustrates subtraction that uses the Subtract(DateTimeOffset) method.")
Public Sub DateTimeOffsetSubtract()
Attribute DateTimeOffsetSubtract.VB_Description = "The following example illustrates subtraction that uses the Subtract(DateTimeOffset) method."
    Dim firstDate As IDateTimeOffset
    Set firstDate = DateTimeOffset.CreateFromDateTimeParts(2018, 10, 25, 18, 0, 0, TimeSpan.Create(-7, 0, 0))
    Dim secondDate As IDateTimeOffset
    Set secondDate = DateTimeOffset.CreateFromDateTimeParts(2018, 10, 25, 18, 0, 0, TimeSpan.Create(-5, 0, 0))
    Dim thirdDate As IDateTimeOffset
    Set thirdDate = DateTimeOffset.CreateFromDateTimeParts(2018, 9, 28, 9, 0, 0, TimeSpan.Create(-7, 0, 0))
    Dim difference As ITimeSpan
    Set difference = firstDate.Subtract(secondDate)
    
    Debug.Print "(" & firstDate.ToString() & ")" & " - " & _
                "(" & secondDate.ToString() & ")" & ": " & _
                difference.days & " days, " & _
                difference.Hours & ":" & _
                format$(difference.Minutes, "00")
            
    Set difference = firstDate.Subtract(thirdDate)
    Debug.Print "(" & firstDate.ToString() & ")" & " - " & _
                "(" & secondDate.ToString() & ")" & ": " & _
                difference.days & " days, " & _
                difference.Hours & ":" & _
                format$(difference.Minutes, "00")
End Sub

' The example produces the following output:
'    (10/25/2018 6:00:00 PM -07:00) - (10/25/2018 6:00:00 PM -05:00): 0 days, 2:00
'    (10/25/2018 6:00:00 PM -07:00) - (9/28/2018 9:00:00 AM -07:00): 27 days, 9:00
