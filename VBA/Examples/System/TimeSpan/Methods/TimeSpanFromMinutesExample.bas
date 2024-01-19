Attribute VB_Name = "TimeSpanFromMinutesExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.fromminutes?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates several TimeSpan objects using the FromMinutes
' method.
''
Public Sub TimeSpanFromMinutes()
    Debug.Print VBString.Unescape( _
        "This example of TimeSpan.FromMinutes( double )\n" + _
        "generates the following output.\n")
    Debug.Print VBString.Format("{0,21}{1,18}", _
                    "FromMinutes", "TimeSpan")
    Debug.Print VBString.Format("{0,21}{1,18}", _
                    "-----------", "--------")
    
    Call GenTimeSpanFromMinutes(0.00001)
    Call GenTimeSpanFromMinutes(0.00002)
    Call GenTimeSpanFromMinutes(0.12345)
    Call GenTimeSpanFromMinutes(1234.56789)
    Call GenTimeSpanFromMinutes(12345678.98765)
    Call GenTimeSpanFromMinutes(0.01666)
    Call GenTimeSpanFromMinutes(1)
    Call GenTimeSpanFromMinutes(60)
    Call GenTimeSpanFromMinutes(1440)
    Call GenTimeSpanFromMinutes(30020.33667)
End Sub

Private Sub GenTimeSpanFromMinutes(ByVal pMinutes As Double)
    ' Create a TimeSpan object and TimeSpan string from
    ' a number of minutes.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.FromMinutes(pMinutes)
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
    
    Debug.Print VBString.Format("{0,21}{1,26}", pMinutes, timeInterval)
End Sub

'/*
'This example of TimeSpan.FromMinutes( double )
'generates the following output.
'
'          FromMinutes          TimeSpan
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

