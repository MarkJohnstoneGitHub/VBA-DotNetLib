Attribute VB_Name = "DateTimeOffsetEqualsExactEg"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified July 21, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.equalsexact?view=net-7.0#examples

Option Explicit

'@Description("The following example illustrates the use of the EqualsExact method to compare similar DateTimeOffset objects.")
Public Sub DateTimeOffsetEqualsExact()
Attribute DateTimeOffsetEqualsExact.VB_Description = "The following example illustrates the use of the EqualsExact method to compare similar DateTimeOffset objects."
    Dim instanceTime As DateTimeOffset
    Set instanceTime = DateTimeOffset.CreateFromDateTimeParts(2007, 10, 31, 0, 0, 0, DateTimeOffset.Now.Offset)
    
    Dim otherTime As DateTimeOffset
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
                
' The example produces the following output:
'       10/31/2007 12:00:00 AM -07:00 = 10/31/2007 12:00:00 AM -07:00: True
'       10/31/2007 12:00:00 AM -07:00 = 10/31/2007 12:00:00 AM -06:00: False
'       10/31/2007 12:00:00 AM -07:00 = 10/31/2007 1:00:00 AM -06:00: False
End Sub
