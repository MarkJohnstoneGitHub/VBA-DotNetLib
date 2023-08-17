Attribute VB_Name = "TimeSpanFromHoursExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.fromhours?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example creates several TimeSpan objects using the FromHours method.")
Public Sub TimeSpanFromHours()
Attribute TimeSpanFromHours.VB_Description = "The following example creates several TimeSpan objects using the FromHours method."
   Debug.Print "FromHours", "TimeSpan"
   Debug.Print "---------", "--------"
   
   GenTimeSpanFromHours 0.0000002
   GenTimeSpanFromHours 0.0000003
   GenTimeSpanFromHours 0.0012345
   GenTimeSpanFromHours 12.3456789
   GenTimeSpanFromHours 123456.7898765
   GenTimeSpanFromHours 0.0002777
   GenTimeSpanFromHours 0.0166666
   GenTimeSpanFromHours 1
   GenTimeSpanFromHours 24
   GenTimeSpanFromHours 500.3389445
End Sub

Private Sub GenTimeSpanFromHours(ByVal Hours As Double)
   ' Create a TimeSpan object and TimeSpan string from
   ' a number of hours.
   Dim interval As ITimeSpan
   Set interval = TimeSpan.FromHours(Hours)
   Dim timeInterval As String
   timeInterval = interval.ToString()
   Debug.Print Hours, timeInterval
End Sub

'/*
'This example of TimeSpan.FromHours( double )
'generates the following output.
'
'            FromHours TimeSpan
'            ---------          --------
'                2E-07          00:00:00.0010000
'                3E-07          00:00:00.0010000
'            0.0012345          00:00:04.4440000
'           12.3456789          12:20:44.4440000
'       123456.7898765     5144.00:47:23.5550000
'            0.0002777          00:00:01
'            0.0166666          00:01:00
'                    1          01:00:00
'                   24        1.00:00:00
'          500.3389445       20.20:20:20.2000000
'*/
