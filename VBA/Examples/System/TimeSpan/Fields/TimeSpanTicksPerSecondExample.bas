Attribute VB_Name = "TimeSpanTicksPerSecondExample"
'@Folder "Examples.System.TimeSpan.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified August 1, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tickspersecond?view=netframework-4.8.1

Option Explicit

'@Description("The following example references and displays the value of the TicksPerSecond field.")
Public Sub TimeSpanTicksPerSecond()
Attribute TimeSpanTicksPerSecond.VB_Description = "The following example references and displays the value of the TicksPerSecond field."
   ' Display the TimeSpan.TicksPerSecond value.
   Debug.Print "TimeSpan.TicksPerSecond : " & Format$(TimeSpan.TicksPerSecond, "#,##0")
End Sub

' Output:
' TimeSpan.TicksPerSecond : 10,000,000
