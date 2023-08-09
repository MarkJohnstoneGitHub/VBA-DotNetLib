Attribute VB_Name = "DateTimeOffsetLocalDateTimeEg"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.localdatetime?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example illustrates several conversions of DateTimeOffset values to local times in the U.S. Pacific Standard Time zone.")
' Note that the last three times are all ambiguous; the property maps all of them to a single date
' and time in the Pacific Standard Time zone.
Public Sub DateTimeOffsetLocalDateTime()
Attribute DateTimeOffsetLocalDateTime.VB_Description = "The following example illustrates several conversions of DateTimeOffset values to local times in the U.S. Pacific Standard Time zone."
    Dim dto As IDateTimeOffset
    
    ' Current time
    Set dto = DateTimeOffset.Now
    Debug.Print dto.LocalDateTime.ToString()
    
    ' Transition to DST in local time zone occurs on 3/11/2007 at 2:00 AM
    Set dto = DateTimeOffset.CreateFromDateTimeParts(2007, 3, 11, 3, 30, 0, TimeSpan.Create(-7, 0, 0))
    Debug.Print dto.LocalDateTime.ToString()
    Set dto = DateTimeOffset.CreateFromDateTimeParts(2007, 3, 11, 2, 30, 0, TimeSpan.Create(-7, 0, 0))
    Debug.Print dto.LocalDateTime.ToString()
    
    ' Invalid time in local time zone
    Set dto = DateTimeOffset.CreateFromDateTimeParts(2007, 3, 11, 2, 30, 0, TimeSpan.Create(-8, 0, 0))
    Debug.Print TimeZoneInfo.Locale.IsInvalidTime(dto.DateTime)
    Debug.Print dto.LocalDateTime.ToString()
    
    ' Transition from DST in local time zone occurs on 11/4/07 at 2:00 AM
    ' This is an ambiguous time
    Set dto = DateTimeOffset.CreateFromDateTimeParts(2007, 11, 4, 1, 30, 0, TimeSpan.Create(-7, 0, 0))
    Debug.Print TimeZoneInfo.Locale.IsAmbiguousTime2(dto)
    Debug.Print dto.LocalDateTime.ToString()
    
    Set dto = DateTimeOffset.CreateFromDateTimeParts(2007, 11, 4, 2, 30, 0, TimeSpan.Create(-7, 0, 0))
    Debug.Print TimeZoneInfo.Locale.IsAmbiguousTime2(dto)
    Debug.Print dto.LocalDateTime.ToString()
    
    Set dto = DateTimeOffset.CreateFromDateTimeParts(2007, 11, 4, 1, 30, 0, TimeSpan.Create(-8, 0, 0))
    Debug.Print TimeZoneInfo.Locale.IsAmbiguousTime2(dto)
    Debug.Print dto.LocalDateTime.ToString()
End Sub

' If run on 3/8/2007 at 4:56 PM, the code produces the following
' output:
'    3/8/2007 4:56:03 PM
'    3/8/2007 4:56:03 PM
'    3/11/2007 3:30:00 AM
'    3/11/2007 1:30:00 AM
'    True
'    3/11/2007 3:30:00 AM
'    True
'    11/4/2007 1:30:00 AM
'    11/4/2007 1:30:00 AM
'    True
'    11/4/2007 1:30:00 AM
