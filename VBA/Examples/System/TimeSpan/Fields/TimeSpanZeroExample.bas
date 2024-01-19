Attribute VB_Name = "TimeSpanZeroExample"
'@Folder "Examples.System.TimeSpan.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.zero?view=netframework-4.8.1

Option Explicit

''
' The following example references and displays the value of the Zero field.
''
Public Sub TimeSpanZero()
   ' Display the zero TimeSpan value.
   Debug.Print "Zero TimeSpan: " & TimeSpan.Zero.ToString
End Sub

' Output:
' Zero TimeSpan: 00:00:00
