Attribute VB_Name = "TimeSpanMaxValueExample"
'@Folder "Examples.System.TimeSpan.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.maxvalue?view=netframework-4.8.1#examples

Option Explicit

''
' The following example references and displays the value of the MaxValue field.
''
Public Sub TimeSpanMaxValue()
   ' Display the maximum  TimeSpan value.
   Debug.Print "Maximum TimeSpan: " & TimeSpan.MaxValue.ToString
End Sub

' Output:
' Maximum TimeSpan: 10675199.02:48:05.4775807
