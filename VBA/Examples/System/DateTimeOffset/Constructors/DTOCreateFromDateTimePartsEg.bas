Attribute VB_Name = "DTOCreateFromDateTimePartsEg"
'@Folder "Examples.System.DateTimeOffset.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified January 9, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.-ctor?view=netframework-4.8.1#system-datetimeoffset-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-timespan)

Option Explicit

''
' The following example instantiates a DateTimeOffset object by using the
' DateTimeOffset.DateTimeOffset(Int32, Int32, Int32, Int32, Int32, Int32, TimeSpan)
' constructor overload.
''
Public Sub DateTimeOffsetCreateFromDateTimeParts()
   Dim specificDate As DotNetLib.DateTime
   Set specificDate = DateTime.CreateFromDateTime(2008, 5, 1, 6, 32, 0)
   Dim offsetDate As DotNetLib.DateTimeOffset
   Set offsetDate = DateTimeOffset.CreateFromDateTimeParts(specificDate.Year, _
                                   specificDate.Month, _
                                   specificDate.Day, _
                                   specificDate.Hour, _
                                   specificDate.Minute, _
                                   specificDate.SECOND, _
                                   TimeSpan.Create(-5, 0, 0))
   Debug.Print VBString.Format("Current time: {0}", offsetDate)
   Debug.Print VBString.Format("Corresponding UTC time: {0}", offsetDate.UtcDateTime)
End Sub

' The code produces the following output:
'    Current time: 5/1/2008 6:32:00 AM -05:00
'    Corresponding UTC time: 5/1/2008 11:32:00 AM

