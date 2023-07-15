Attribute VB_Name = "TimeSpanTicksPerDayExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Fields")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.ticksperday?view=netframework-4.8.1

Option Explicit

'@Description("The following example references and displays the value of the TicksPerDay field.")
Public Sub TimeSpanTicksPerDay()
Attribute TimeSpanTicksPerDay.VB_Description = "The following example references and displays the value of the TicksPerDay field."
   ' Display the TimeSpan.TicksPerDay.
   Debug.Print "TimeSpan.TicksPerDay : " & TimeSpan.TicksPerDay
   
' Output:
' TimeSpan.TicksPerDay : 864000000000
End Sub
