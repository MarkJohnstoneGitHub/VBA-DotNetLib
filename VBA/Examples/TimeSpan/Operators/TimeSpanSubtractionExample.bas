Attribute VB_Name = "TimeSpanSubtractionExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Operators")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified July 17, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.op_subtraction?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the TimeSpan subtraction operator to calculate the total length of the weekly work day.")
' It also uses the TimeSpan addition operator to compute the total time of the daily breaks before
' using it in a subtraction operation to compute the total actual daily working time.
Public Sub TimeSpanSubtraction()
Attribute TimeSpanSubtraction.VB_Description = "The following example uses the TimeSpan subtraction operator to calculate the total length of the weekly work day."
   Dim startWork As TimeSpan
   Set startWork = TimeSpan.Create(8, 0, 0)
   Dim endWork As TimeSpan
   Set endWork = TimeSpan.Create(18, 30, 0)
   Dim lunchBreak As TimeSpan
   Set lunchBreak = TimeSpan.Create(1, 0, 0)
   Dim breaks As TimeSpan
   Set breaks = TimeSpan.Create(0, 30, 0)
   
   Debug.Print "Length of work day: " & TimeSpan.Subtraction(endWork, startWork).ToString()
   Debug.Print "Actual time worked: " & _
               TimeSpan.Subtraction(TimeSpan.Subtraction(endWork, startWork), TimeSpan.Addition(lunchBreak, breaks)).ToString()
               
' The example displays the following output:
'     Length of work day: 10:30:00
'     Actual time worked: 09:00:00
End Sub
