Attribute VB_Name = "TimeSpanFieldsExample"
'@Folder "Examples.System.TimeSpan.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 16, 2024
'@LastModified January 16, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.maxvalue?view=netframework-4.8.1#examples

Option Explicit

''
' Example of the TimeSpan fields.
''
Public Sub TimeSpanFieldsExample()
    Const numberFmt As String = "{0,-22}{1,18:N0}"
    Const timeFmt As String = "{0,-22}{1,26}"
    
    Debug.Print "This example of the fields of the TimeSpan class" & _
                VBString.Unescape("\ngenerates the following output.\n")
    Debug.Print VBString.Format(numberFmt, "Field", "Value")
    Debug.Print VBString.Format(numberFmt, "-----", "-----")
    
    ' Display the maximum, minimum, and zero TimeSpan values.
    Debug.Print VBString.Format(timeFmt, "Maximum TimeSpan", _
            Align(TimeSpan.MaxValue))
    Debug.Print VBString.Format(timeFmt, "Minimum TimeSpan", _
            Align(TimeSpan.MinValue))
    Debug.Print VBString.Format(timeFmt, "Zero TimeSpan", _
            Align(TimeSpan.Zero))
    Debug.Print
    
    ' Display the ticks-per-time-unit fields.
    Debug.Print VBString.Format(numberFmt, "Ticks per day", _
                                TimeSpan.TicksPerDay)
    Debug.Print VBString.Format(numberFmt, "Ticks per hour", _
                                TimeSpan.TicksPerHour)
    Debug.Print VBString.Format(numberFmt, "Ticks per minute", _
                                TimeSpan.TicksPerMinute)
    Debug.Print VBString.Format(numberFmt, "Ticks per second", _
                                TimeSpan.TicksPerSecond)
    Debug.Print VBString.Format(numberFmt, "Ticks per millisecond", _
                                TimeSpan.TicksPerMillisecond)
End Sub

Private Function Align(ByVal pInterval As DotNetLib.TimeSpan) As String
    Dim intervalStr As String
    intervalStr = pInterval.ToString()
    Dim pointIndex As Long
    pointIndex = InStr(intervalStr, ":")
    pointIndex = InStr(pointIndex, intervalStr, ".")
    If (pointIndex = 0) Then
        intervalStr = intervalStr & "        "
    End If
    Align = intervalStr
End Function

'/*
'This example of the fields of the TimeSpan class
'generates the following output.
'
'Field                              Value
'-----                              -----
'Maximum TimeSpan       10675199.02:48:05.4775807
'Minimum TimeSpan      -10675199.02:48:05.4775808
'Zero TimeSpan                   00:00:00
'
'Ticks per day            864,000,000,000
'Ticks per hour            36,000,000,000
'Ticks per minute             600,000,000
'Ticks per second              10,000,000
'Ticks per millisecond             10,000
'*/

