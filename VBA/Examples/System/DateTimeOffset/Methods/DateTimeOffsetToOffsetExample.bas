Attribute VB_Name = "DateTimeOffsetToOffsetExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tooffset?view=netframework-4.8.1#examples

Option Explicit

Private sourceTime As DotNetLib.DateTimeOffset

''
' The following example illustrates how to use the ToOffset method to convert a
' DateTimeOffset object to a DateTimeOffset object with a different offset.
''
Public Sub DateTimeOffsetToOffset()
    Dim targetTime As DotNetLib.DateTimeOffset
    Set sourceTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 9, 30, 0, TimeSpan.Create(-5, 0, 0))
    
    ' Convert to same time (return sourceTime unchanged)
    Set targetTime = sourceTime.ToOffset(TimeSpan.Create(-5, 0, 0))
    Call ShowDateAndTimeInfo(targetTime)
    
    ' Convert to UTC (0 offset)
    Set targetTime = sourceTime.ToOffset(TimeSpan.Zero)
    Call ShowDateAndTimeInfo(targetTime)
    
    ' Convert to 8 hours behind UTC
    Set targetTime = sourceTime.ToOffset(TimeSpan.Create(-8, 0, 0))
    Call ShowDateAndTimeInfo(targetTime)

    ' Convert to 3 hours ahead of UTC
    Set targetTime = sourceTime.ToOffset(TimeSpan.Create(3, 0, 0))
    Call ShowDateAndTimeInfo(targetTime)
End Sub

Private Sub ShowDateAndTimeInfo(ByVal newTime As DotNetLib.DateTimeOffset)
    Debug.Print VBString.Format("{0} converts to {1}", sourceTime, newTime)
    Debug.Print VBString.Format("{0} and {1} are equal: {2}", _
                                sourceTime, newTime, sourceTime.Equals(newTime))
    Debug.Print VBString.Format("{0} and {1} are identical: {2}", _
                                sourceTime, newTime, _
                                sourceTime.EqualsExact(newTime))
    Debug.Print
End Sub

' The example displays the following output:
'    9/1/2007 9:30:00 AM -05:00 converts to 9/1/2007 9:30:00 AM -05:00
'    9/1/2007 9:30:00 AM -05:00 and 9/1/2007 9:30:00 AM -05:00 are equal: True
'    9/1/2007 9:30:00 AM -05:00 and 9/1/2007 9:30:00 AM -05:00 are identical: True
'
'    9/1/2007 9:30:00 AM -05:00 converts to 9/1/2007 2:30:00 PM +00:00
'    9/1/2007 9:30:00 AM -05:00 and 9/1/2007 2:30:00 PM +00:00 are equal: True
'    9/1/2007 9:30:00 AM -05:00 and 9/1/2007 2:30:00 PM +00:00 are identical: False
'
'    9/1/2007 9:30:00 AM -05:00 converts to 9/1/2007 6:30:00 AM -08:00
'    9/1/2007 9:30:00 AM -05:00 and 9/1/2007 6:30:00 AM -08:00 are equal: True
'    9/1/2007 9:30:00 AM -05:00 and 9/1/2007 6:30:00 AM -08:00 are identical: False
'
'    9/1/2007 9:30:00 AM -05:00 converts to 9/1/2007 5:30:00 PM +03:00
'    9/1/2007 9:30:00 AM -05:00 and 9/1/2007 5:30:00 PM +03:00 are equal: True
'    9/1/2007 9:30:00 AM -05:00 and 9/1/2007 5:30:00 PM +03:00 are identical: False

