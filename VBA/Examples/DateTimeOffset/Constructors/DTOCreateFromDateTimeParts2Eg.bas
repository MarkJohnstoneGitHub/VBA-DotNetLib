Attribute VB_Name = "DTOCreateFromDateTimeParts2Eg"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified July 18, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.-ctor?view=netframework-4.8.1#system-datetimeoffset-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-timespan)

Option Explicit

'@Description("The following example instantiates a DateTimeOffset object by using the DateTimeOffset.DateTimeOffset(Int32, Int32, Int32, Int32, Int32, Int32, Int32, TimeSpan) constructor overload.")
Public Sub DateTimeOffsetCreateFromDateTimeParts2()
Attribute DateTimeOffsetCreateFromDateTimeParts2.VB_Description = "The following example instantiates a DateTimeOffset object by using the DateTimeOffset.DateTimeOffset(Int32, Int32, Int32, Int32, Int32, Int32, Int32, TimeSpan) constructor overload."
   Dim fmt As String
   fmt = "dd MMM yyyy HH:mm:ss"
   Dim thisDate As DateTime
   Set thisDate = DateTime.CreateFromDateTime(2007, 6, 12, 19, 0, 14, 16)
   Dim offsetDate As DateTimeOffset
   
   Set offsetDate = DateTimeOffset.CreateFromDateTimeParts2(thisDate.Year, _
                                   thisDate.Month, _
                                   thisDate.Day, _
                                   thisDate.Hour, _
                                   thisDate.Minute, _
                                   thisDate.Second, _
                                   thisDate.Millisecond, _
                                   TimeSpan.Create(-5, 0, 0))
   Debug.Print "Current time: " & offsetDate.ToString2(fmt) & ":" & offsetDate.Millisecond
   
' The code produces the following output:
'    Current time: 12 Jun 2007 19:00:14:16
End Sub
