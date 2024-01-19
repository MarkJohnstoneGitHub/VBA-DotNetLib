Attribute VB_Name = "TimeSpanCreateExample"
'@Folder "Examples.System.TimeSpan.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 15, 2023
'@LastModified January 14, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.-ctor?view=netframework-4.8.1#system-timespan-ctor(system-int32-system-int32-system-int32)

Option Explicit

''
' The following example creates several TimeSpan objects using the constructor
' overload that initializes a TimeSpan to a specified number of hours, minutes,
' and seconds.
''
Public Sub TimeSpanCreate()
    Debug.Print "This example of the TimeSpan( int, int, int )" & VBA.vbNewLine & _
                 "constructor generates the following output." & VBA.vbNewLine
    Debug.Print VBString.Format("{0,-37}{1,16}", "Constructor", "Value")
    Debug.Print VBString.Format("{0,-37}{1,16}", "-----------", "-----")
    
    Call CreateTimeSpan(10, 20, 30)
    Call CreateTimeSpan(-10, 20, 30)
    Call CreateTimeSpan(0, 0, 37230)
    Call CreateTimeSpan(1000, 2000, 3000)
    Call CreateTimeSpan(1000, -2000, -3000)
    Call CreateTimeSpan(999999, 999999, 999999)
End Sub

Private Sub CreateTimeSpan(ByVal pHours As Long, ByVal pMinutes As Long, ByVal pSeconds As Long)
    Dim elapsedTime As DotNetLib.TimeSpan
    Set elapsedTime = TimeSpan.Create(pHours, pMinutes, pSeconds)
    
    ' Format the constructor for display.
    Dim ctor As String
    ctor = VBString.Format("TimeSpan( {0}, {1}, {2} )", _
                            pHours, pMinutes, pSeconds)

    ' Display the constructor and its value.
    Debug.Print VBString.Format("{0,-37}{1,16}", _
                                ctor, elapsedTime.ToString())
End Sub

'/*
'This example of the TimeSpan( int, int, int )
'constructor generates the following output.
'
'Constructor value
'-----------                                     -----
'TimeSpan( 10, 20, 30 )                       10:20:30
'TimeSpan( -10, 20, 30 )                     -09:39:30
'TimeSpan( 0, 0, 37230 )                      10:20:30
'TimeSpan( 1000, 2000, 3000 )              43.02:10:00
'TimeSpan( 1000, -2000, -3000 )            40.05:50:00
'TimeSpan( 999999, 999999, 999999 )     42372.15:25:39
'*/

