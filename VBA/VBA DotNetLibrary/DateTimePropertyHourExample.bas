Attribute VB_Name = "DateTimePropertyHourExample"
'@Folder("Examples.DateTime")

'https://learn.microsoft.com/en-us/dotnet/api/system.datetime.hour?view=netframework-4.8.1#examples

Option Explicit

Public Sub DateTimePropertyHour()
    Dim date1 As DateTime
    Set date1 = DateTime.CreateFromDateTime(2008, 4, 1, 18, 53, 0)
    Debug.Print date1.ToString2("%h")       ' Displays 6
    Debug.Print date1.ToString2("h tt")     ' Displays 6 PM
End Sub


