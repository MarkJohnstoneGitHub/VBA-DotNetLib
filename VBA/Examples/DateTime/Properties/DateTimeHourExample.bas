Attribute VB_Name = "DateTimeHourExample"
'Rubberduck annotations
'@Folder "VBADotNetLib.Examples.DateTime.Properties"

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.hour?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the Hour property.")
Public Sub DateTimePropertyHour()
Attribute DateTimePropertyHour.VB_Description = "The following example demonstrates the Hour property."
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTime(2008, 4, 1, 18, 53, 0)
    Debug.Print date1.ToString2("%h")       ' Displays 6
    Debug.Print date1.ToString2("h tt")     ' Displays 6 PM
End Sub
