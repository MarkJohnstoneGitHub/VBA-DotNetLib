Attribute VB_Name = "TimeSpanFromMinutesExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.fromminutes?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example creates several TimeSpan objects using the FromMinutes method.")
Public Sub TimeSpanFromMinutes()
Attribute TimeSpanFromMinutes.VB_Description = "The following example creates several TimeSpan objects using the FromMinutes method."
   Debug.Print "FromMinutes", "TimeSpan"
   Debug.Print "-----------", "--------"

   GenTimeSpanFromMinutes 0.00001
   GenTimeSpanFromMinutes 0.00002
   GenTimeSpanFromMinutes 0.12345
   GenTimeSpanFromMinutes 1234.56789
   GenTimeSpanFromMinutes 12345678.98765
   GenTimeSpanFromMinutes 0.01666
   GenTimeSpanFromMinutes 1
   GenTimeSpanFromMinutes 60
   GenTimeSpanFromMinutes 1440
   GenTimeSpanFromMinutes 30020.33667
End Sub

Private Sub GenTimeSpanFromMinutes(ByVal Minutes As Double)
   ' Create a TimeSpan object and TimeSpan string from
   ' a number of minutes.
   Dim interval As ITimeSpan
   Set interval = TimeSpan.FromMinutes(Minutes)
   Dim timeInterval As String
   timeInterval = interval.ToString()
   Debug.Print Minutes, timeInterval
End Sub

'/*
'This example of TimeSpan.FromMinutes( double )
'generates the following output.
'
'          FromMinutes TimeSpan
'          -----------          --------
'                1E-05          00:00:00.0010000
'                2E-05          00:00:00.0010000
'              0.12345          00:00:07.4070000
'           1234.56789          20:34:34.0730000
'       12345678.98765     8573.09:18:59.2590000
'              0.01666          00:00:01
'                    1          00:01:00
'                   60          01:00:00
'                 1440        1.00:00:00
'          30020.33667       20.20:20:20.2000000
'*/
