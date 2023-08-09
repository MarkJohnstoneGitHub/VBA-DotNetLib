Attribute VB_Name = "TimeSpanFromTicksExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.fromticks?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example creates several TimeSpan objects using the FromTicks method.")
Public Sub TimeSpanFromTicks()
Attribute TimeSpanFromTicks.VB_Description = "The following example creates several TimeSpan objects using the FromTicks method."
   Debug.Print "FromTicks", "TimeSpan"
   Debug.Print "---------", "--------"
   
   GenTimeSpanFromTicks 1
   GenTimeSpanFromTicks 12345
   GenTimeSpanFromTicks 123456789
   GenTimeSpanFromTicks 1234567898765#
   GenTimeSpanFromTicks "12345678987654321"
   GenTimeSpanFromTicks 10000000
   GenTimeSpanFromTicks 600000000
   GenTimeSpanFromTicks 36000000000#
   GenTimeSpanFromTicks 864000000000#
   GenTimeSpanFromTicks 18012202000000#
End Sub

Private Sub GenTimeSpanFromTicks(ByVal Ticks As LongLong)
   ' Create a TimeSpan object and TimeSpan string from
   ' a number of seconds.
   Dim interval As ITimeSpan
   Set interval = TimeSpan.FromTicks(Ticks)
   Dim timeInterval As String
   timeInterval = interval.ToString()
   Debug.Print Ticks, timeInterval
End Sub

'/*
'This example of TimeSpan.FromTicks( long )
'generates the following output.
'
'            FromTicks TimeSpan
'            ---------          --------
'                    1          00:00:00.0000001
'                12345          00:00:00.0012345
'            123456789          00:00:12.3456789
'        1234567898765        1.10:17:36.7898765
'    12345678987654321    14288.23:31:38.7654321
'             10000000          00:00:01
'            600000000          00:01:00
'          36000000000          01:00:00
'         864000000000        1.00:00:00
'       18012202000000       20.20:20:20.2000000
'*/
