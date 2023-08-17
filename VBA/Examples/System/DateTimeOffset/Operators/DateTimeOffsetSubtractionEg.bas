Attribute VB_Name = "DateTimeOffsetSubtractionEg"
'@Folder "Examples.System.DateTimeOffset.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.op_subtraction?view=netframework-4.8.1#system-datetimeoffset-op-subtraction(system-datetimeoffset-system-datetimeoffset)

Option Explicit

'@Description("The Subtraction method defines the subtraction operation for DateTimeOffset objects.")
Public Sub DateTimeOffsetSubtraction()
Attribute DateTimeOffsetSubtraction.VB_Description = "The Subtraction method defines the subtraction operation for DateTimeOffset objects."
    Dim firstDate As IDateTimeOffset
    Set firstDate = DateTimeOffset.CreateFromDateTimeParts(2008, 3, 25, 18, 0, 0, TimeSpan.Create(-7, 0, 0))
    Dim secondDate As IDateTimeOffset
    Set secondDate = DateTimeOffset.CreateFromDateTimeParts(2008, 3, 25, 18, 0, 0, TimeSpan.Create(-5, 0, 0))
    Dim thirdDate As IDateTimeOffset
    Set thirdDate = DateTimeOffset.CreateFromDateTimeParts(2008, 2, 28, 9, 0, 0, TimeSpan.Create(-7, 0, 0))
    Dim difference As ITimeSpan
    
    Set difference = DateTimeOffset.Subtraction(firstDate, secondDate)
    Debug.Print "(" & firstDate.ToString() & ")" & " - " & _
                "(" & secondDate.ToString() & ")" & ": " & _
                difference.Days & " days, " & _
                difference.Hours & ":" & _
                format$(difference.Minutes, "00")
                
    Set difference = DateTimeOffset.Subtraction(firstDate, thirdDate)
    Debug.Print "(" & firstDate.ToString() & ")" & " - " & _
                "(" & secondDate.ToString() & ")" & ": " & _
                difference.Days & " days, " & _
                difference.Hours & ":" & _
                format$(difference.Minutes, "00")
End Sub

' The example produces the following output:
'    (3/25/2008 6:00:00 PM -07:00) - (3/25/2008 6:00:00 PM -05:00): 0 days, 2:00
'    (3/25/2008 6:00:00 PM -07:00) - (3/25/2008 6:00:00 PM -05:00): 26 days, 9:00
