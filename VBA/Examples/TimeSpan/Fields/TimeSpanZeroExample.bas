Attribute VB_Name = "TimeSpanZeroExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Fields")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.zero?view=netframework-4.8.1

Option Explicit

'@Description("The following example references and displays the value of the Zero field.")
Public Sub TimeSpanZero()
Attribute TimeSpanZero.VB_Description = "The following example references and displays the value of the Zero field."
   ' Display the zero TimeSpan value.
   Debug.Print "Zero TimeSpan: " & TimeSpan.Zero.ToString
   
' Output:
' Zero TimeSpan: 00:00:00
End Sub
