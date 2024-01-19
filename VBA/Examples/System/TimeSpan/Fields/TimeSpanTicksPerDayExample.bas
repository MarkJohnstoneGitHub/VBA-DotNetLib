Attribute VB_Name = "TimeSpanTicksPerDayExample"
'@Folder "Examples.System.TimeSpan.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.ticksperday?view=netframework-4.8.1

Option Explicit

''
' The following example references and displays the value of the TicksPerDay field.
''
Public Sub TimeSpanTicksPerDay()
   ' Display the TimeSpan.TicksPerDay.
   Debug.Print "TimeSpan.TicksPerDay : " & Format$(TimeSpan.TicksPerDay, "#,##0")
End Sub

' Output:
' TimeSpan.TicksPerDay : 864,000,000,000
