Attribute VB_Name = "TimeSpanFromMillisecondsExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.frommilliseconds?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates several TimeSpan objects by using the
' FromMilliseconds method.
''
Public Sub TimeSpanFromMilliseconds()
    Debug.Print VBString.Unescape( _
        "This example of TimeSpan.FromMilliseconds( " + _
        "double )\ngenerates the following output.\n")
    Debug.Print VBString.Format("{0,21}{1,18}", _
        "FromMilliseconds", "TimeSpan")
    Debug.Print VBString.Format("{0,21}{1,18}", _
        "----------------", "--------")
   
    Call GenTimeSpanFromMillisec(1)
    Call GenTimeSpanFromMillisec(1.5)
    Call GenTimeSpanFromMillisec(12345.6)
    Call GenTimeSpanFromMillisec(123456789.8)
    Call GenTimeSpanFromMillisec(1234567898765.4)
    Call GenTimeSpanFromMillisec(1000)
    Call GenTimeSpanFromMillisec(60000)
    Call GenTimeSpanFromMillisec(3600000)
    Call GenTimeSpanFromMillisec(86400000)
    Call GenTimeSpanFromMillisec(1801220200)
End Sub

Private Sub GenTimeSpanFromMillisec(ByVal millisec As Double)
    ' Create a TimeSpan object and TimeSpan string from
    ' a number of milliseconds.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.FromMilliseconds(millisec)
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
        
    Debug.Print VBString.Format("{0,21}{1,26}", millisec, timeInterval)
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

