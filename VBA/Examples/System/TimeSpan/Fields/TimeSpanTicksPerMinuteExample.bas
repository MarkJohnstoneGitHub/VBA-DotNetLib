Attribute VB_Name = "TimeSpanTicksPerMinuteExample"
'@Folder "Examples.System.TimeSpan.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.ticksperminute?view=netframework-4.8.1

Option Explicit

''
' The following example references and displays the value of the TicksPerMinute
' field.
''
Public Sub TimeSpanTicksPerMinute()
   ' Display the TimeSpan.TicksPerMinute value.
   Debug.Print "TimeSpan.TicksPerMinute : " & Format$(TimeSpan.TicksPerMinute, "#,##0")
End Sub

' Output:
' TimeSpan.TicksPerMinute : 600,000,000
