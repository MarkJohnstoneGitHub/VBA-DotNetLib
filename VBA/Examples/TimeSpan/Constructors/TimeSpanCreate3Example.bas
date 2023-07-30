Attribute VB_Name = "TimeSpanCreate3Example"
'@Folder("VBADotNetLib.Examples.TimeSpan.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.-ctor?view=netframework-4.8.1#system-timespan-ctor(system-int32-system-int32-system-int32-system-int32-system-int32)

Option Explicit

'@Description("The following example creates several TimeSpan objects using the constructor overload that initializes a TimeSpan to a specified number of days, hours, minutes, seconds, and milliseconds.")
Public Sub TimeSpanCreate3()
Attribute TimeSpanCreate3.VB_Description = "The following example creates several TimeSpan objects using the constructor overload that initializes a TimeSpan to a specified number of days, hours, minutes, seconds, and milliseconds."
   Debug.Print "This example of the TimeSpan( int, int, int, int, int ) " & VBA.vbNewLine & "constructor generates the following output." & VBA.vbNewLine
   Debug.Print "Constructor", "Value"
   Debug.Print "-----------", "-----"
   CreateTimeSpan 10, 20, 30, 40, 50
   CreateTimeSpan -10, 20, 30, 40, 50
   CreateTimeSpan 0, 0, 0, 0, 937840050
   CreateTimeSpan 1111, 2222, 3333, 4444, 5555
   CreateTimeSpan 1111, -2222, -3333, -4444, -5555
   CreateTimeSpan 99999, 99999, 99999, 99999, 99999
End Sub

'/*
'This example of the TimeSpan( int, int, int, int, int )
'constructor generates the following output.
'
'Constructor value
'-----------                                                -----
'TimeSpan( 10, 20, 30, 40, 50 )                       10.20:30:40.0500000
'TimeSpan( -10, 20, 30, 40, 50 )                      -9.03:29:19.9500000
'TimeSpan( 0, 0, 0, 0, 937840050 )                    10.20:30:40.0500000
'TimeSpan( 1111, 2222, 3333, 4444, 5555 )           1205.22:47:09.5550000
'TimeSpan( 1111, -2222, -3333, -4444, -5555 )       1016.01:12:50.4450000
'TimeSpan( 99999, 99999, 99999, 99999, 99999 )    104236.05:27:18.9990000
'*/

Private Sub CreateTimeSpan(ByVal Days As Long, ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long, ByVal millisec As Long)
   Dim elapsedTime As ITimeSpan
   Set elapsedTime = TimeSpan.Create3(Days, Hours, Minutes, Seconds, millisec)

   ' Format the constructor for display.
   Dim ctor As String
   ctor = "TimeSpan( " & Days & ", " & Hours & ", " & Minutes & ", " & Seconds & ", " & millisec & " )"
   
   ' Display the constructor and its value.
   Debug.Print ctor, elapsedTime.ToString()
End Sub
