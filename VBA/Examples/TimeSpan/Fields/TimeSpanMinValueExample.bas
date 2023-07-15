Attribute VB_Name = "TimeSpanMinValueExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Fields")
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.minvalue?view=netframework-4.8.1

Option Explicit

'@Description("The following example references and displays the value of the MinValue field.")
Public Sub TimeSpanMinValue()
Attribute TimeSpanMinValue.VB_Description = "The following example references and displays the value of the MinValue field."
   ' Display the minimum TimeSpan value.
   Debug.Print "Minimum TimeSpan: " & TimeSpan.MinValue.ToString
   
' Output:
' Minimum TimeSpan: -10675199.02:48:05.4775808
End Sub
