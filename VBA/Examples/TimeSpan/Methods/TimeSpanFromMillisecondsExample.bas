Attribute VB_Name = "TimeSpanFromMillisecondsExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.frommilliseconds?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example creates several TimeSpan objects by using the FromMilliseconds method.")
Public Sub TimeSpanFromMilliseconds()
Attribute TimeSpanFromMilliseconds.VB_Description = "The following example creates several TimeSpan objects by using the FromMilliseconds method."
   Debug.Print "FromMilliseconds", "TimeSpan"
   Debug.Print "----------------", "--------"
   
   GenTimeSpanFromMillisec 1
   GenTimeSpanFromMillisec 1.5
   GenTimeSpanFromMillisec 12345.6
   GenTimeSpanFromMillisec 123456789.8
   GenTimeSpanFromMillisec 1234567898765.4
   GenTimeSpanFromMillisec 1000
   GenTimeSpanFromMillisec 60000
   GenTimeSpanFromMillisec 3600000
   GenTimeSpanFromMillisec 86400000
   GenTimeSpanFromMillisec 1801220200
End Sub

Private Sub GenTimeSpanFromMillisec(ByVal millisec As Double)
   ' Create a TimeSpan object and TimeSpan string from
   ' a number of milliseconds.
   Dim interval As ITimeSpan
   Set interval = TimeSpan.FromMilliseconds(millisec)
   Dim timeInterval As String
   timeInterval = interval.ToString()
   Debug.Print millisec, timeInterval
End Sub

'/*
'This example of TimeSpan.FromMilliseconds( double )
'generates the following output.
'
'     FromMilliseconds TimeSpan
'     ----------------          --------
'                    1          00:00:00.0010000
'                  1.5          00:00:00.0020000
'              12345.6          00:00:12.3460000
'          123456789.8        1.10:17:36.7900000
'      1234567898765.4    14288.23:31:38.7650000
'                 1000          00:00:01
'                60000          00:01:00
'              3600000          01:00:00
'             86400000        1.00:00:00
'           1801220200       20.20:20:20.2000000
'*/
