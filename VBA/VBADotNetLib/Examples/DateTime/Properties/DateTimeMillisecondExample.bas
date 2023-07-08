Attribute VB_Name = "DateTimeMillisecondExample"
'Rubberduck annotations
'@Folder "VBADotNetLib.Examples.DateTime.Properties"

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 09, 2023

'@DotNetReference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.millisecond?view=net-7.0#examples

Option Explicit

Public Sub DateTimePropertyMillisecond()
    Dim date1 As DateTime
    Set date1 = DateTime.CreateFromDateTime(2008, 1, 1, 0, 30, 45, 125)
    Debug.Print "Milliseconds: " & date1.ToString2("fff")

    Dim date2 As DateTime
    Set date2 = DateTime.CreateFromDateTime(2008, 1, 1, 0, 30, 45, 125)
    Debug.Print "Milliseconds: " & date2.ToString2("o")
    ' Displays the following output to the console:
    '     Date: 2008-01-01T00:30:45.1250000
End Sub

