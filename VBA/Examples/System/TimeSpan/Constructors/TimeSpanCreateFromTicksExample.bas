Attribute VB_Name = "TimeSpanCreateFromTicksExample"
'@Folder "Examples.System.TimeSpan.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 15, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.-ctor?view=netframework-4.8.1#system-timespan-ctor(system-int64)

Option Explicit

''
' The following example creates several TimeSpan objects using the constructor
' overload that initializes a TimeSpan to a specified number of ticks.
''
Public Sub TimeSpanCreateFromTicks()
    Debug.Print VBString.Unescape( _
        "This example of the TimeSpan( long ) constructor " + _
        "\ngenerates the following output.\n")
    Debug.Print VBString.Format("{0,-33}{1,16}", "Constructor", "Value")
    Debug.Print VBString.Format("{0,-33}{1,16}", "-----------", "-----")
   
    Call CreateTimeSpan(1)
    Call CreateTimeSpan(999999)
    Call CreateTimeSpan(-1000000000000#)
    Call CreateTimeSpan(18012202000000#)
    Call CreateTimeSpan("999999999999999999")
    Call CreateTimeSpan("1000000000000000000")
End Sub

Private Sub CreateTimeSpan(ByVal pTicks As LongLong)
    Dim elapsedTime As DotNetLib.TimeSpan
    Set elapsedTime = TimeSpan.CreateFromTicks(pTicks)
    
    ' Format the constructor for display.
    Dim ctor As String
    ctor = VBString.Format("TimeSpan( {0} )", pTicks)

    ' Pad the end of a TimeSpan string with spaces if
    ' it does not contain milliseconds.
    Dim pointIndex As Long
    Dim elapsedStr As String
    elapsedStr = elapsedTime.ToString()
    pointIndex = InStr(elapsedStr, ":")
    pointIndex = InStr(pointIndex, elapsedStr, ".")
    If (pointIndex = 0) Then
        elapsedStr = elapsedStr & "        "
    End If
    
    ' Display the constructor and its value.
    Debug.Print VBString.Format("{0,-33}{1,24}", ctor, elapsedStr)
End Sub

'This example of the TimeSpan( long ) constructor
'generates the following output.
'
'Constructor                                 Value
'-----------                                 -----
'TimeSpan( 1 )                            00:00:00.0000001
'TimeSpan( 999999 )                       00:00:00.0999999
'TimeSpan( -1000000000000 )            -1.03:46:40
'TimeSpan( 18012202000000 )            20.20:20:20.2000000
'TimeSpan( 999999999999999999 )   1157407.09:46:39.9999999
'TimeSpan( 1000000000000000000 )  1157407.09:46:40

