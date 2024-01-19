Attribute VB_Name = "TimeSpanFromHoursExample"
'@Folder "Examples.System.DateTime"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 17, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.fromhours?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates several TimeSpan objects using the FromHours method.
''
Public Sub TimeSpanFromHours()
    Debug.Print VBString.Unescape( _
        "This example of TimeSpan.FromHours( double )\n" + _
        "generates the following output.\n")
    Debug.Print VBString.Format("{0,21}{1,18}", _
        "FromHours", "TimeSpan")
    Debug.Print VBString.Format("{0,21}{1,18}", _
        "---------", "--------")
    
    Call GenTimeSpanFromHours(0.0000002)
    Call GenTimeSpanFromHours(0.0000003)
    Call GenTimeSpanFromHours(0.0012345)
    Call GenTimeSpanFromHours(12.3456789)
    Call GenTimeSpanFromHours(123456.7898765)
    Call GenTimeSpanFromHours(0.0002777)
    Call GenTimeSpanFromHours(0.0166666)
    Call GenTimeSpanFromHours(1)
    Call GenTimeSpanFromHours(24)
    Call GenTimeSpanFromHours(500.3389445)
End Sub

Private Sub GenTimeSpanFromHours(ByVal pHours As Double)
    ' Create a TimeSpan object and TimeSpan string from
    ' a number of hours.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.FromHours(pHours)
    Dim timeInterval As String
    timeInterval = interval.ToString()
    
    ' Pad the end of the TimeSpan string with spaces if it
    ' does not contain milliseconds.
    Dim pIndex As Long
    pIndex = InStr(timeInterval, ":")
    pIndex = InStr(pIndex, timeInterval, ".")
    If (pIndex = 0) Then
        timeInterval = timeInterval & "        "
    End If
        
    Debug.Print VBString.Format("{0,21}{1,26}", pHours, timeInterval)
End Sub

'/*
'This example of TimeSpan.FromHours( double )
'generates the following output.
'
'            FromHours          TimeSpan
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


