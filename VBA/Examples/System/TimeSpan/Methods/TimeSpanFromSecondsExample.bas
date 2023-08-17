Attribute VB_Name = "TimeSpanFromSecondsExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.fromseconds?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example creates several TimeSpan objects using the FromSeconds method.")
Public Sub TimeSpanFromSeconds()
Attribute TimeSpanFromSeconds.VB_Description = "The following example creates several TimeSpan objects using the FromSeconds method."
   Debug.Print "FromSeconds", "TimeSpan"
   Debug.Print "-----------", "--------"
   
   GenTimeSpanFromSeconds 0.001
   GenTimeSpanFromSeconds 0.0015
   GenTimeSpanFromSeconds 12.3456
   GenTimeSpanFromSeconds 123456.7898
   GenTimeSpanFromSeconds 1234567898.7654
   GenTimeSpanFromSeconds 1
   GenTimeSpanFromSeconds 60
   GenTimeSpanFromSeconds 3600
   GenTimeSpanFromSeconds 86400
   GenTimeSpanFromSeconds 1801220.2
End Sub

Private Sub GenTimeSpanFromSeconds(ByVal Seconds As Double)
   ' Create a TimeSpan object and TimeSpan string from
   ' a number of seconds.
   Dim interval As ITimeSpan
   Set interval = TimeSpan.FromSeconds(Seconds)
   Dim timeInterval As String
   timeInterval = interval.ToString()
   Debug.Print Seconds, timeInterval
End Sub

'/*
'This example of TimeSpan.FromSeconds( double )
'generates the following output.
'
'          FromSeconds TimeSpan
'          -----------          --------
'                0.001          00:00:00.0010000
'               0.0015          00:00:00.0020000
'              12.3456          00:00:12.3460000
'          123456.7898        1.10:17:36.7900000
'      1234567898.7654    14288.23:31:38.7650000
'                    1          00:00:01
'                   60          00:01:00
'                 3600          01:00:00
'                86400        1.00:00:00
'            1801220.2       20.20:20:20.2000000
'*/
