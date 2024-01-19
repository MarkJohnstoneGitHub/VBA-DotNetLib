Attribute VB_Name = "DateTimeOffsetToUniversalTimeEg"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.touniversaltime?view=netframework-4.8.1#examples

Option Explicit

''
' The following example calls the ToUniversalTime method to convert a local
' time and several other times to Coordinated Universal Time (UTC).
''
Public Sub DateTimeOffsetToUniversalTime()
    Dim localTime As DotNetLib.DateTimeOffset
    Dim otherTime As DotNetLib.DateTimeOffset
    Dim universalTime As DotNetLib.DateTimeOffset
    
    ' Define local time in local time zone
    Set localTime = DateTimeOffset.CreateFromDateTime(DateTime.CreateFromDateTime(2007, 6, 15, 12, 0, 0))
    Debug.Print VBString.Format("Local time: {0}", localTime)
    Debug.Print
    
    ' Convert local time to offset 0 and assign to otherTime
    Set otherTime = localTime.ToOffset(TimeSpan.Zero)
    Debug.Print VBString.Format("Other time: {0}", otherTime)
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                localTime, otherTime, _
                                localTime.Equals(otherTime))
    Debug.Print VBString.Format("{0} exactly equals {1}: {2}", _
                                localTime, otherTime, _
                                localTime.EqualsExact(otherTime))
    Debug.Print
    
    ' Convert other time to UTC
    Set universalTime = localTime.ToUniversalTime()
    Debug.Print VBString.Format("Universal time: {0}", universalTime)
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                otherTime, universalTime, _
                                universalTime.Equals(otherTime))
    Debug.Print VBString.Format("{0} exactly equals {1}: {2}", _
                                otherTime, universalTime, _
                                universalTime.EqualsExact(otherTime))
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

