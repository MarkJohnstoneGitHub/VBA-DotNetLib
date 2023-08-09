Attribute VB_Name = "DTOCreateFromDateTimePartsEg"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.-ctor?view=netframework-4.8.1#system-datetimeoffset-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-timespan)

Option Explicit

'@Description("The following example instantiates a DateTimeOffset object by using the DateTimeOffset.DateTimeOffset(Int32, Int32, Int32, Int32, Int32, Int32, TimeSpan) constructor overload.")
Public Sub DateTimeOffsetCreateFromDateTimeParts()
Attribute DateTimeOffsetCreateFromDateTimeParts.VB_Description = "The following example instantiates a DateTimeOffset object by using the DateTimeOffset.DateTimeOffset(Int32, Int32, Int32, Int32, Int32, Int32, TimeSpan) constructor overload."
   Dim specificDate As IDateTime
   Set specificDate = DateTime.CreateFromDateTime(2008, 5, 1, 6, 32, 0)
   Dim offsetDate As IDateTimeOffset
   
   Set offsetDate = DateTimeOffset.CreateFromDateTimeParts(specificDate.Year, _
                                   specificDate.Month, _
                                   specificDate.Day, _
                                   specificDate.Hour, _
                                   specificDate.Minute, _
                                   specificDate.Second, _
                                   TimeSpan.Create(-5, 0, 0))
   Debug.Print "Current time: " & offsetDate.ToString()
   Debug.Print "Corresponding UTC time: " & offsetDate.UtcDateTime.ToString()
End Sub

' The code produces the following output:
'    Current time: 5/1/2008 6:32:00 AM -05:00
'    Corresponding UTC time: 5/1/2008 11:32:00 AM
