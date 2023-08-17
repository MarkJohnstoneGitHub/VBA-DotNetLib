Attribute VB_Name = "DateTimeCreateFromDateTimeEg"
'@Folder "Examples.System.DateTime.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 3, 2023
'@LastModified August 3, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32)

Option Explicit

'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32) constructor to instantiate a DateTime value.")
' Initializes a new instance of the DateTime structure to the specified
' year, month, day, hour, minute, and second.
Public Sub DateTimeCreateFromDateTime()
Attribute DateTimeCreateFromDateTime.VB_Description = "The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32) constructor to instantiate a DateTime value."
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 0)
    Debug.Print date1.ToString()
End Sub

' The example displays the following output, in this case for en-us culture:
'      8/18/2010 4:32:00 PM
