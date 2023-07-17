Attribute VB_Name = "DateTimeOffsetCreateFromDateTime2Example"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified July 18, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.-ctor?view=netframework-4.8.1#system-datetimeoffset-ctor(system-datetime-system-timespan)

Option Explicit

'@Description("The following example shows how to initialize a DateTimeOffset object with a date and time and the offset of the local time zone when that time zone is not known in advance.")
Public Sub DateTimeOffsetCreateFromDateTime2()
Attribute DateTimeOffsetCreateFromDateTime2.VB_Description = "The following example shows how to initialize a DateTimeOffset object with a date and time and the offset of the local time zone when that time zone is not known in advance."
   Dim localTime As DateTime
   Set localTime = DateTime.CreateFromDateTime(2007, 7, 12, 6, 32, 0)
   Dim dateAndOffset As DateTimeOffset
   Set dateAndOffset = DateTimeOffset.CreateFromDateTime2(localTime, TimeZoneInfo.Locale.GetUtcOffset(localTime))
   Debug.Print dateAndOffset.ToString()
   
' The code produces the following output:
'    7/12/2007 6:32:00 AM -07:00
End Sub
