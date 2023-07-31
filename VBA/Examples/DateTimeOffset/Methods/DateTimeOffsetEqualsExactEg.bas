Attribute VB_Name = "DateTimeOffsetEqualsExactEg"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.equalsexact?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example illustrates the use of the EqualsExact method to compare similar DateTimeOffset objects.")
Public Sub DateTimeOffsetEqualsExact()
Attribute DateTimeOffsetEqualsExact.VB_Description = "The following example illustrates the use of the EqualsExact method to compare similar DateTimeOffset objects."
    Dim instanceTime As IDateTimeOffset
    Set instanceTime = DateTimeOffset.CreateFromDateTimeParts(2007, 10, 31, 0, 0, 0, DateTimeOffset.Now.Offset)
    
    Dim otherTime As IDateTimeOffset
    Set otherTime = instanceTime
    Debug.Print instanceTime.ToString() & " = " & _
                otherTime.ToString() & ": " & _
                instanceTime.EqualsExact(otherTime)

    Set otherTime = DateTimeOffset.CreateFromDateTime2(instanceTime.DateTime, TimeSpan.FromHours(instanceTime.Offset.Hours + 1))
    Debug.Print instanceTime.ToString() & " = " & _
                otherTime.ToString() & ": " & _
                instanceTime.EqualsExact(otherTime)
                
    Set otherTime = DateTimeOffset.CreateFromDateTime2(DateTime.Addition(instanceTime.DateTime, TimeSpan.FromHours(1)), TimeSpan.FromHours(instanceTime.Offset.Hours + 1))
    Debug.Print instanceTime.ToString() & " = " & _
                otherTime.ToString() & ": " & _
                instanceTime.EqualsExact(otherTime)
End Sub

' The example produces the following output:
'       10/31/2007 12:00:00 AM -07:00 = 10/31/2007 12:00:00 AM -07:00: True
'       10/31/2007 12:00:00 AM -07:00 = 10/31/2007 12:00:00 AM -06:00: False
'       10/31/2007 12:00:00 AM -07:00 = 10/31/2007 1:00:00 AM -06:00: False
