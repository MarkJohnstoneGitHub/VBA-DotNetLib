Attribute VB_Name = "TimeSpanTicksPerMillisecondEg"
'@Folder "Examples.System.TimeSpan.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified August 1, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tickspermillisecond?view=netframework-4.8.1

Option Explicit

'@Description("The following example references and displays the value of the TicksPerMillisecond field.")
Public Sub TimeSpanTicksPerMillisecond()
Attribute TimeSpanTicksPerMillisecond.VB_Description = "The following example references and displays the value of the TicksPerMillisecond field."
   ' Display the TimeSpan.TicksPerMillisecond value.
   Debug.Print "TimeSpan.TicksPerMillisecond : " & format$(TimeSpan.TicksPerMillisecond, "#,##0")
End Sub

' Output:
' TimeSpan.TicksPerMillisecond : 10,000