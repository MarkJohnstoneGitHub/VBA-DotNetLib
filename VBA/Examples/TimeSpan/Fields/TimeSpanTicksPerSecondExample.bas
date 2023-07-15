Attribute VB_Name = "TimeSpanTicksPerSecondExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Fields")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tickspersecond?view=netframework-4.8.1

Option Explicit

'@Description("The following example references and displays the value of the TicksPerSecond field.")
Public Sub TimeSpanTicksPerSecond()
Attribute TimeSpanTicksPerSecond.VB_Description = "The following example references and displays the value of the TicksPerSecond field."
   ' Display the TimeSpan.TicksPerSecond value.
   Debug.Print "TimeSpan.TicksPerSecond : " & TimeSpan.TicksPerSecond
   
' Output:
' TimeSpan.TicksPerSecond : 10000000
End Sub
