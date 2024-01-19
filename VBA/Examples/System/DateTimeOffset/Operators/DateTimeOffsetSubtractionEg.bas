Attribute VB_Name = "DateTimeOffsetSubtractionEg"
'@Folder "Examples.System.DateTimeOffset.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.op_subtraction?view=netframework-4.8.1#system-datetimeoffset-op-subtraction(system-datetimeoffset-system-datetimeoffset)

Option Explicit

''
' The Subtraction method defines the subtraction operation for DateTimeOffset objects.
'  It enables code such as the following:
''
Public Sub DateTimeOffsetSubtraction()
    Dim firstDate As DotNetLib.DateTimeOffset
    Set firstDate = DateTimeOffset.CreateFromDateTimeParts(2008, 3, 25, 18, 0, 0, TimeSpan.Create(-7, 0, 0))
    Dim secondDate As DotNetLib.DateTimeOffset
    Set secondDate = DateTimeOffset.CreateFromDateTimeParts(2008, 3, 25, 18, 0, 0, TimeSpan.Create(-5, 0, 0))
    Dim thirdDate As DotNetLib.DateTimeOffset
    Set thirdDate = DateTimeOffset.CreateFromDateTimeParts(2008, 2, 28, 9, 0, 0, TimeSpan.Create(-7, 0, 0))
    Dim difference As DotNetLib.TimeSpan
    
    Set difference = DateTimeOffset.Subtraction(firstDate, secondDate)
    Debug.Print VBString.Format("({0}) - ({1}): {2} days, {3}:{4:d2}", _
                                firstDate.ToString(), _
                                secondDate.ToString(), _
                                difference.Days, _
                                difference.Hours, _
                                difference.Minutes)
                
    Set difference = DateTimeOffset.Subtraction(firstDate, thirdDate)
    Debug.Print VBString.Format("({0}) - ({1}): {2} days, {3}:{4:d2}", _
                                firstDate.ToString(), _
                                thirdDate.ToString(), _
                                difference.Days, _
                                difference.Hours, _
                                difference.Minutes)
End Sub

' The example produces the following output:
'    (3/25/2008 6:00:00 PM -07:00) - (3/25/2008 6:00:00 PM -05:00): 0 days, 2:00
'    (3/25/2008 6:00:00 PM -07:00) - (3/25/2008 6:00:00 PM -05:00): 26 days, 9:00

