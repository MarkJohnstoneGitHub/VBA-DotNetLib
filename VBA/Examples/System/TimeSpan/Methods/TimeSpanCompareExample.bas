Attribute VB_Name = "TimeSpanCompareExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.compare?view=netframework-4.8.1#examples

Option Explicit

'@Description("'The following example uses the Compare method to compare several TimeSpan objects with a TimeSpan object whose value is a 2-hour time interval.")
Public Sub TimeSpanCompare()
Attribute TimeSpanCompare.VB_Description = "'The following example uses the Compare method to compare several TimeSpan objects with a TimeSpan object whose value is a 2-hour time interval."
   ' Define a time interval equal to two hours.
   Dim baseInterval As ITimeSpan
   Set baseInterval = TimeSpan.Create(2, 0, 0)
   
   ' Define an array of time intervals to compare with
   ' the base interval.
   Dim spans() As ITimeSpan
   Objects.ToArray spans, _
                  TimeSpan.FromSeconds(-2.5), _
                  TimeSpan.FromMinutes(20), _
                  TimeSpan.FromHours(1), _
                  TimeSpan.FromMinutes(90), _
                  baseInterval, _
                  TimeSpan.FromDays(0.5), _
                  TimeSpan.FromDays(1)
                  
   ' Compare the time intervals.
   Dim varSpan As Variant
   For Each varSpan In spans
      Dim span As ITimeSpan
      Set span = varSpan
      Dim result As Long
      result = TimeSpan.Compare(baseInterval, span)
      
      Debug.Print baseInterval.ToString() _
            & " " & _
            IIf(result = 1, ">", IIf(result = 0, "=", "<")) _
            & " " _
            & span.ToString() _
            & " (Compare returns " _
            & result _
            & ")"
   Next
End Sub

' The example displays the following output:
'       02:00:00 > -00:00:02.5000000 (Compare returns 1)
'       02:00:00 > 00:20:00 (Compare returns 1)
'       02:00:00 > 01:00:00 (Compare returns 1)
'       02:00:00 > 01:30:00 (Compare returns 1)
'       02:00:00 = 02:00:00 (Compare returns 0)
'       02:00:00 < 12:00:00 (Compare returns -1)
'       02:00:00 < 1.00:00:00 (Compare returns -1)
