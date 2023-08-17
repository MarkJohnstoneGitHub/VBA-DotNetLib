Attribute VB_Name = "DateTimeOffsetToUniversalTimeEg"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.touniversaltime?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example calls the ToUniversalTime method to convert a local time and several other times to Coordinated Universal Time (UTC).")
Public Sub DateTimeOffsetToUniversalTime()
Attribute DateTimeOffsetToUniversalTime.VB_Description = "The following example calls the ToUniversalTime method to convert a local time and several other times to Coordinated Universal Time (UTC)."
    Dim localTime As IDateTimeOffset
    Dim otherTime As IDateTimeOffset
    Dim universalTime As IDateTimeOffset
    
    ' Define local time in local time zone
    Set localTime = DateTimeOffset.CreateFromDateTime(DateTime.CreateFromDateTime(2007, 6, 15, 12, 0, 0))
    Debug.Print "Local time: " & localTime.ToString()
    Debug.Print
    
    ' Convert local time to offset 0 and assign to otherTime
    Set otherTime = localTime.ToOffset(TimeSpan.Zero)
    Debug.Print "Other time: " & otherTime.ToString()
    Debug.Print localTime.ToString() & " = " & _
                otherTime.ToString() & ": " & _
                localTime.Equals(otherTime)
                
    Debug.Print localTime.ToString() & " exactly equals " & _
                otherTime.ToString() & ": " & _
                localTime.EqualsExact(otherTime)
    Debug.Print
    
    ' Convert other time to UTC
    Set universalTime = localTime.ToUniversalTime()
    Debug.Print "Universal time: " & universalTime.ToString()
    Debug.Print otherTime.ToString() & " = " & _
                universalTime.ToString() & ": " & _
                universalTime.Equals(otherTime)
                
    Debug.Print otherTime.ToString() & " exactly equals " & _
                universalTime.ToString() & ": " & _
                universalTime.EqualsExact(otherTime)
    Debug.Print
End Sub

' The example produces the following output to the console:
'    Local time: 6/15/2007 12:00:00 PM -07:00
'
'    Other time: 6/15/2007 7:00:00 PM +00:00
'    6/15/2007 12:00:00 PM -07:00 = 6/15/2007 7:00:00 PM +00:00: True
'    6/15/2007 12:00:00 PM -07:00 exactly equals 6/15/2007 7:00:00 PM +00:00: False
'
'    Universal time: 6/15/2007 7:00:00 PM +00:00
'    6/15/2007 7:00:00 PM +00:00 = 6/15/2007 7:00:00 PM +00:00: True
'    6/15/2007 7:00:00 PM +00:00 exactly equals 6/15/2007 7:00:00 PM +00:00: True
