Attribute VB_Name = "TimeSpanToStringExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified July 17, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tostring?view=netframework-4.8.1#system-timespan-tostring

Option Explicit

'@Description("The following example displays the strings returned by calling the ToString method with a number of TimeSpan values.")
Public Sub TimeSpanToString()
Attribute TimeSpanToString.VB_Description = "The following example displays the strings returned by calling the ToString method with a number of TimeSpan values."
   Dim span As TimeSpan
   ' Initialize a time span to zero.
   Set span = TimeSpan.Zero
   Debug.Print span.ToString()

   ' Initialize a time span to 14 days.
   Set span = TimeSpan.Create3(-14, 0, 0, 0, 0)
   Debug.Print span.ToString()
   
   ' Initialize a time span to 1:02:03.
   Set span = TimeSpan.Create(1, 2, 3)
   Debug.Print span.ToString()
   
   ' // Initialize a time span to 250 milliseconds.
   Set span = TimeSpan.Create3(0, 0, 0, 0, 250)
   Debug.Print span.ToString
   
   ' // Initialize a time span to 99 days, 23 hours, 59 minutes, and 59.999 seconds.
   Set span = TimeSpan.Create3(99, 23, 59, 59, 999)
   Debug.Print span.ToString()
   
   ' Initialize a time span to 3 hours.
   Set span = TimeSpan.Create(3, 0, 0)
   Debug.Print span.ToString
   
   ' Initialize a timespan to 25 milliseconds.
   Set span = TimeSpan.Create3(0, 0, 0, 0, 25)
   Debug.Print span.ToString

' The example displays the following output:
'        00:00:00
'        -14.00:00:00
'        01:02:03
'        00:00:00.2500000
'        99.23:59:59.9990000
'        03:00:00
'        00:00:00.0250000
End Sub
