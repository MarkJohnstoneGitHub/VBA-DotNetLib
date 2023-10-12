Attribute VB_Name = "DateTimeAdditionExample"
'@Folder "Examples.System.DateTime.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.op_addition?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the addition operator.")
Public Sub DateTimeAddition()
Attribute DateTimeAddition.VB_Description = "The following example demonstrates the addition operator."
    Dim dTime As IDateTime
    Set dTime = DateTime.CreateFromDate(1980, 8, 5)
    
    ' tSpan is 17 days, 4 hours, 2 minutes and 1 second.
    Dim tSpan As ITimeSpan
    Set tSpan = TimeSpan.Create2(17, 4, 2, 1)
    
    ' Result gets 8/22/1980 4:02:01 AM.
    Dim Result As IDateTime
    Set Result = DateTime.Addition(dTime, tSpan)
    Debug.Print Result.ToString
End Sub
