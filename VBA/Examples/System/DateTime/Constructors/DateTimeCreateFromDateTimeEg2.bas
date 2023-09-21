Attribute VB_Name = "DateTimeCreateFromDateTimeEg2"
'@Folder "Examples.System.DateTime.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 3, 2023
'@LastModified August 3, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-int32)

'Note: DateTime.CreateFromDateTime millisecond parameter is optional

Option Explicit

'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, Int32) constructor to instantiate a DateTime value.")
' Initializes a new instance of the DateTime structure to the specified
' year, month, day, hour, minute, second, and millisecond.
Public Sub DateTimeCreateFromDateTimeMillisecond()
Attribute DateTimeCreateFromDateTimeMillisecond.VB_Description = "The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, Int32) constructor to instantiate a DateTime value."
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 18, 500)
    Debug.Print date1.ToString2("M/dd/yyyy h:mm:ss.fff tt")
End Sub

' The example displays the following output, in this case for en-us culture:
' 8/18/2010 4:32:18.500 PM
