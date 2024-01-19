Attribute VB_Name = "DateTimeOffsetEqualsExactEg"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.equalsexact?view=netframework-4.8.1#examples

Option Explicit

''
' The following example illustrates the use of the EqualsExact method to compare
' similar DateTimeOffset objects.
''
Public Sub DateTimeOffsetEqualsExact()
    Dim instanceTime As DotNetLib.DateTimeOffset
    Set instanceTime = DateTimeOffset.CreateFromDateTimeParts(2007, 10, 31, 0, 0, 0, _
                                        DateTimeOffset.Now.Offset)
    
    Dim otherTime As DotNetLib.DateTimeOffset
    Set otherTime = instanceTime
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                instanceTime, otherTime, _
                                instanceTime.EqualsExact(otherTime))

    Set otherTime = DateTimeOffset.CreateFromDateTime2(instanceTime.DateTime, _
                                    TimeSpan.FromHours(instanceTime.Offset.Hours + 1))
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                instanceTime, otherTime, _
                                instanceTime.EqualsExact(otherTime))
                
    Set otherTime = DateTimeOffset.CreateFromDateTime2(DateTime.Addition(instanceTime.DateTime, TimeSpan.FromHours(1)), _
                                    TimeSpan.FromHours(instanceTime.Offset.Hours + 1))
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                instanceTime, otherTime, _
                                instanceTime.EqualsExact(otherTime))
End Sub

' The example produces the following output:
'       10/31/2007 12:00:00 AM -07:00 = 10/31/2007 12:00:00 AM -07:00: True
'       10/31/2007 12:00:00 AM -07:00 = 10/31/2007 12:00:00 AM -06:00: False
'       10/31/2007 12:00:00 AM -07:00 = 10/31/2007 1:00:00 AM -06:00: False

