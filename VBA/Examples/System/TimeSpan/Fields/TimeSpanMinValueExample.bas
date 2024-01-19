Attribute VB_Name = "TimeSpanMinValueExample"
'@Folder "Examples.System.TimeSpan.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.minvalue?view=netframework-4.8.1

Option Explicit

''
' The following example references and displays the value of the MinValue field.
''
Public Sub TimeSpanMinValue()
   ' Display the minimum TimeSpan value.
   Debug.Print "Minimum TimeSpan: " & TimeSpan.MinValue.ToString
End Sub

' Output:
' Minimum TimeSpan: -10675199.02:48:05.4775808
