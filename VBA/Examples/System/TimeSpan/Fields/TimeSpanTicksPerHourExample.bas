Attribute VB_Name = "TimeSpanTicksPerHourExample"
'@Folder "Examples.System.TimeSpan.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.ticksperhour?view=netframework-4.8.1

Option Explicit

''
' The following example references and displays the value of the TicksPerHour field.
''
Public Sub TimeSpanTicksPerHour()
   ' Display the TimeSpan.TicksPerHour value.
   Debug.Print "TimeSpan.TicksPerHour : " & Format$(TimeSpan.TicksPerHour, "#,##0")
End Sub

' Output:
' TimeSpan.TicksPerHour : 36,000,000,000
