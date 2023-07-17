Attribute VB_Name = "TimeSpanAdditionExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Operators")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified July 17, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.op_addition?view=netframework-4.8.1

Option Explicit

'@Description("The Addition method defines the addition operator for TimeSpan values.")
Public Sub TimeSpanAddition()
Attribute TimeSpanAddition.VB_Description = "The Addition method defines the addition operator for TimeSpan values."
   Dim time1 As TimeSpan
   Set time1 = TimeSpan.Create2(1, 0, 0, 0)     ' TimeSpan equivalent to 1 day.
   Dim time2 As TimeSpan
   Set time2 = TimeSpan.Create(12, 0, 0)        ' TimeSpan equivalent to 1/2 day.
   Dim time3 As TimeSpan
   Set time3 = TimeSpan.Addition(time1, time2)  ' Add the two time spans.
   
   Debug.Print "   " & time1.ToString()
   Debug.Print " +   " & time2.ToString()
   Debug.Print "   " & "__________"
   Debug.Print "   " & time3.ToString()
   
' The example displays the following output:
'           1.00:00:00
'        +    12:00:00
'          __________
'           1.12:00:00
End Sub
