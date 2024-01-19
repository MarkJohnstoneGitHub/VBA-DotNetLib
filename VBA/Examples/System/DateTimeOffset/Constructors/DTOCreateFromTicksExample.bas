Attribute VB_Name = "DTOCreateFromTicksExample"
'@Folder "Examples.System.DateTimeOffset.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified January 9, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.-ctor?view=netframework-4.8.1#system-datetimeoffset-ctor(system-int64-system-timespan)

Option Explicit

''
' The following example initializes a DateTimeOffset object by using the number
' of ticks in an arbitrary date (in this case, July 16, 2007, at 1:32 PM) with
' an offset of -5.
''
Public Sub DateTimeOffsetCreateFromTicks()
   Dim dateWithoutOffset As DotNetLib.DateTime
   Set dateWithoutOffset = DateTime.CreateFromDateTime(2007, 7, 16, 13, 32, 0)
   Dim timeFromTicks As DotNetLib.DateTimeOffset
   Set timeFromTicks = DateTimeOffset.CreateFromTicks(dateWithoutOffset.Ticks, TimeSpan.Create(-5, 0, 0))
   Debug.Print timeFromTicks.ToString()
End Sub

' The code produces the following output:
'    7/16/2007 1:32:00 PM -05:00
