Attribute VB_Name = "TimeSpanTicksPerMinuteExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Fields")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.ticksperminute?view=netframework-4.8.1

Option Explicit

'@Description("The following example references and displays the value of the TicksPerMinute field.")
Public Sub TimeSpanTicksPerMinute()
Attribute TimeSpanTicksPerMinute.VB_Description = "The following example references and displays the value of the TicksPerMinute field."
   ' Display the TimeSpan.TicksPerMinute value.
   Debug.Print "TimeSpan.TicksPerMinute : " & TimeSpan.TicksPerMinute
   
' Output:
' TimeSpan.TicksPerMinute : 600000000
End Sub