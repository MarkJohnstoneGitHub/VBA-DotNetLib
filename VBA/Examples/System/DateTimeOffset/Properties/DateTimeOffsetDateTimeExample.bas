Attribute VB_Name = "DateTimeOffsetDateTimeExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.datetime?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example illustrates the use of the DateTime property to convert the time returned by the Now and UtcNow properties to DateTime values.")
Public Sub DateTimeOffsetDateTime()
Attribute DateTimeOffsetDateTime.VB_Description = "The following example illustrates the use of the DateTime property to convert the time returned by the Now and UtcNow properties to DateTime values."
   Dim offsetDate As IDateTimeOffset
   Dim regularDate As IDateTime
   
   Set offsetDate = DateTimeOffset.Now
   Set regularDate = offsetDate.DateTime
   Debug.Print offsetDate.ToString() & " converts to " & regularDate.ToString() & ", Kind " & DateTimeKindHelper.ToString(regularDate.Kind)
   
   Set offsetDate = DateTimeOffset.UtcNow
   Set regularDate = offsetDate.DateTime
   Debug.Print offsetDate.ToString() & " converts to " & regularDate.ToString() & ", Kind " & DateTimeKindHelper.ToString(regularDate.Kind)
End Sub

' If run on 3/6/2007 at 17:11, produces the following output:
'
'   3/6/2007 5:11:22 PM -08:00 converts to 3/6/2007 5:11:22 PM, Kind Unspecified.
'   3/7/2007 1:11:22 AM +00:00 converts to 3/7/2007 1:11:22 AM, Kind Unspecified.
