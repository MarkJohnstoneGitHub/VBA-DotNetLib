Attribute VB_Name = "TimeSpanCompareToExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.compareto?view=netframework-4.8.1#system-timespan-compareto(system-timespan)

Option Explicit

'@Description("Compares this instance to a specified TimeSpan object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the TimeSpan object.")
Public Sub TimeSpanCompareTo()
Attribute TimeSpanCompareTo.VB_Description = "Compares this instance to a specified TimeSpan object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the TimeSpan object."
   Dim tsX As ITimeSpan
   Set tsX = TimeSpan.Create2(11, 22, 33, 44)
   
   Dim tsFirst As ITimeSpan
   Dim tsSecond As ITimeSpan
   Set tsFirst = tsX
   Set tsSecond = tsX
   
   Dim result As Long
   result = tsFirst.CompareTo(tsSecond)
   
   Debug.Print tsFirst.ToString() _
         & " " _
         & IIf(result = 1, ">", IIf(result = 0, "=", "<")) _
         & " " _
         & tsSecond.ToString() _
         & " (Compare returns " _
         & result _
         & ")"
End Sub

' Output:
' 11.22:33:44 = 11.22:33:44 (Compare returns 0)
