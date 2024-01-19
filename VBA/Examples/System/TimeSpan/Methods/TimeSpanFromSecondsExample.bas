Attribute VB_Name = "TimeSpanFromSecondsExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.fromseconds?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates several TimeSpan objects using the FromSeconds
' method.
''
Public Sub TimeSpanFromSeconds()
    Debug.Print VBString.Unescape( _
        "This example of TimeSpan.FromSeconds( double )\n" + _
        "generates the following output.\n")
    Debug.Print VBString.Format("{0,21}{1,18}", _
        "FromSeconds", "TimeSpan")
    Debug.Print VBString.Format("{0,21}{1,18}", _
        "-----------", "--------")
    
    Call GenTimeSpanFromSeconds(0.001)
    Call GenTimeSpanFromSeconds(0.0015)
    Call GenTimeSpanFromSeconds(12.3456)
    Call GenTimeSpanFromSeconds(123456.7898)
    Call GenTimeSpanFromSeconds(1234567898.7654)
    Call GenTimeSpanFromSeconds(1)
    Call GenTimeSpanFromSeconds(60)
    Call GenTimeSpanFromSeconds(3600)
    Call GenTimeSpanFromSeconds(86400)
    Call GenTimeSpanFromSeconds(1801220.2)
End Sub

Private Sub GenTimeSpanFromSeconds(ByVal Seconds As Double)
    ' Create a TimeSpan object and TimeSpan string from
    ' a number of seconds.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.FromSeconds(Seconds)
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
    
    Debug.Print VBString.Format("{0,21}{1,26}", Seconds, timeInterval)
End Sub

'/*
'This example of TimeSpan.FromSeconds( double )
'generates the following output.
'
'          FromSeconds          TimeSpan
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

