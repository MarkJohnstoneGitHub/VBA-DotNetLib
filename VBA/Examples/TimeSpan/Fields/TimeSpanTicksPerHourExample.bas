Attribute VB_Name = "TimeSpanTicksPerHourExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Fields")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.ticksperhour?view=netframework-4.8.1

Option Explicit

'@Description("The following example references and displays the value of the TicksPerHour field.")
Public Sub TimeSpanTicksPerHour()
Attribute TimeSpanTicksPerHour.VB_Description = "The following example references and displays the value of the TicksPerHour field."
   ' Display the TimeSpan.TicksPerHour value.
   Debug.Print "TimeSpan.TicksPerHour : " & TimeSpan.TicksPerHour
   
' Output:
' TimeSpan.TicksPerHour : 36000000000
End Sub
