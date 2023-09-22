Attribute VB_Name = "DateTimeCreateFromDateExample"
'@Folder "Examples.System.DateTime.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 3, 2023
'@LastModified August 3, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32)

Option Explicit

'@Description("The following example uses the DateTime(Int32, Int32, Int32) constructor to instantiate a DateTime value. The example also illustrates that this overload creates a DateTime value whose time component equals midnight (or 0:00).")
' Initializes a new instance of the DateTime structure to the specified
' year, month, and day.
Public Sub DateTimeCreateFromDate()
Attribute DateTimeCreateFromDate.VB_Description = "The following example uses the DateTime(Int32, Int32, Int32) constructor to instantiate a DateTime value. The example also illustrates that this overload creates a DateTime value whose time component equals midnight (or 0:00)."
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDate(2010, 8, 18)
    Debug.Print date1.ToString()
End Sub

' The example displays the following output:
'      8/18/2010 12:00:00 AM


'@TODO Implement Console to open the command window etc.
Public Sub DateTimeCreateFromDateConsole()
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDate(2010, 8, 18)
    Console.WriteLine date1.ToString
End Sub

