Attribute VB_Name = "DateTimeAdditionExample"
'@Folder "Examples.System.DateTime.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.op_addition?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the addition operator.")
Public Sub DateTimeAddition()
Attribute DateTimeAddition.VB_Description = "The following example demonstrates the addition operator."
    Dim dTime As DotNetLib.DateTime
    Set dTime = DateTime.CreateFromDate(1980, 8, 5)
    
    ' tSpan is 17 days, 4 hours, 2 minutes and 1 second.
    Dim tSpan As DotNetLib.TimeSpan
    Set tSpan = TimeSpan.Create2(17, 4, 2, 1)
    
    ' Result gets 8/22/1980 4:02:01 AM.
    Dim result As DotNetLib.DateTime
    Set result = DateTime.Addition(dTime, tSpan)
    Debug.Print result.ToString
End Sub
