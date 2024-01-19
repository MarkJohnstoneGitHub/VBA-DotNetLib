Attribute VB_Name = "DateTimeMillisecondExample"
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.millisecond?view=netframework-4.8.1#remarks

Option Explicit

''
' The following example demonstrates the Millisecond property.
''
Public Sub DateTimePropertyMillisecond()
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime(2008, 1, 1, 0, 30, 45, 125)
    Debug.Print VBString.Format("Milliseconds: {0:fff}", _
                                date1)  ' displays Milliseconds: 125
                  
    Dim date2 As DotNetLib.DateTime
    Set date2 = DateTime.CreateFromDateTime(2008, 1, 1, 0, 30, 45, 125)
    Debug.Print VBString.Format("Date: {0:o}", _
                                date2);
End Sub

' Displays the following output:
'    Milliseconds: 125
'    Date: 2008-01-01T00:30:45.1250000


