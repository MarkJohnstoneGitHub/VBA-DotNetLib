Attribute VB_Name = "TimeSpanFromTicksExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.fromticks?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates several TimeSpan objects using the FromTicks method.
''
Public Sub TimeSpanFromTicks()
    Debug.Print VBString.Unescape( _
        "This example of TimeSpan.FromTicks( long )\n" + _
        "generates the following output.\n")
    Debug.Print VBString.Format("{0,21}{1,18}", _
        "FromTicks", "TimeSpan")
    Debug.Print VBString.Format("{0,21}{1,18}", _
        "---------", "--------")
        
    Call GenTimeSpanFromTicks(1)
    Call GenTimeSpanFromTicks(12345)
    Call GenTimeSpanFromTicks(123456789)
    Call GenTimeSpanFromTicks(1234567898765#)
    Call GenTimeSpanFromTicks("12345678987654321")
    Call GenTimeSpanFromTicks(10000000)
    Call GenTimeSpanFromTicks(600000000)
    Call GenTimeSpanFromTicks(36000000000#)
    Call GenTimeSpanFromTicks(864000000000#)
    Call GenTimeSpanFromTicks(18012202000000#)
End Sub

Private Sub GenTimeSpanFromTicks(ByVal pTicks As LongLong)
    ' Create a TimeSpan object and TimeSpan string from
    ' a number of seconds.
    Dim interval As DotNetLib.TimeSpan
    Set interval = TimeSpan.FromTicks(pTicks)
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
    
    Debug.Print VBString.Format("{0,21}{1,26}", pTicks, timeInterval)
End Sub

'/*
'This example of TimeSpan.FromTicks( long )
'generates the following output.
'
'            FromTicks          TimeSpan
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

