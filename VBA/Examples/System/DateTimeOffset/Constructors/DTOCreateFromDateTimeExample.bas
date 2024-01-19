Attribute VB_Name = "DTOCreateFromDateTimeExample"
'@Folder "Examples.System.DateTimeOffset.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified January 8, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.-ctor?view=netframework-4.8.1#system-datetimeoffset-ctor(system-datetime)

Option Explicit

'@Description("The following example illustrates how the value of the DateTime.Kind property of the dateTime parameter affects the date and time value that is returned by this constructor.")
Public Sub DateTimeOffsetCreateFromDateTime()
Attribute DateTimeOffsetCreateFromDateTime.VB_Description = "The following example illustrates how the value of the DateTime.Kind property of the dateTime parameter affects the date and time value that is returned by this constructor."
   Dim localNow As DotNetLib.DateTime
   Set localNow = DateTime.Now
   Dim localOffset As DotNetLib.DateTimeOffset
   Set localOffset = DateTimeOffset.CreateFromDateTime(localNow)
   Debug.Print localOffset.ToString()
   
   Dim pvtUtcNow As DotNetLib.DateTime
   Set pvtUtcNow = DateTime.UtcNow
   Dim utcOffset As DotNetLib.DateTimeOffset
   Set utcOffset = DateTimeOffset.CreateFromDateTime(pvtUtcNow)
   Debug.Print utcOffset.ToString()
   
   Dim unspecifiedNow As DotNetLib.DateTime
   Set unspecifiedNow = DateTime.SpecifyKind(DateTime.Now, DateTimeKind.DateTimeKind_Unspecified)
   Dim unspecifiedOffset As DotNetLib.DateTimeOffset
   Set unspecifiedOffset = DateTimeOffset.CreateFromDateTime(unspecifiedNow)
   Debug.Print unspecifiedOffset.ToString()
End Sub

' The code produces the following output if run on Feb. 23, 2007, on
' a system 8 hours earlier than UTC:
'   2/23/2007 4:21:58 PM -08:00
'   2/24/2007 12:21:58 AM +00:00
'   2/23/2007 4:21:58 PM -08:00
