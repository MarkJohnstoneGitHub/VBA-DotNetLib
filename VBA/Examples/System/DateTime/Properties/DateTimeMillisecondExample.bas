Attribute VB_Name = "DateTimeMillisecondExample"
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.millisecond?view=netframework-4.8.1#examples

Option Explicit

Public Sub DateTimePropertyMillisecond()
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTime(2008, 1, 1, 0, 30, 45, 125)
    Debug.Print "Milliseconds: " & date1.ToString2("fff")

    Dim date2 As IDateTime
    Set date2 = DateTime.CreateFromDateTime(2008, 1, 1, 0, 30, 45, 125)
    Debug.Print "Date: " & date2.ToString2("o")
End Sub

' Displays the following output:
'    Milliseconds: 125
'    Date: 2008-01-01T00:30:45.1250000
