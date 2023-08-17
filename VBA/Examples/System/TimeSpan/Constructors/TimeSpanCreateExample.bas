Attribute VB_Name = "TimeSpanCreateExample"
'@Folder "Examples.System.TimeSpan.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 15, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.-ctor?view=netframework-4.8.1#system-timespan-ctor(system-int32-system-int32-system-int32)

Option Explicit

'@Description("The following example creates several TimeSpan objects using the constructor overload that initializes a TimeSpan to a specified number of hours, minutes, and seconds.")
Public Sub TimeSpanCreate()
Attribute TimeSpanCreate.VB_Description = "The following example creates several TimeSpan objects using the constructor overload that initializes a TimeSpan to a specified number of hours, minutes, and seconds."
   Debug.Print "This example of the TimeSpan( int, int, int )" & VBA.vbNewLine & "constructor generates the following output." & VBA.vbNewLine
   Debug.Print "Constructor", "Value"
   Debug.Print "-----------", "-----"
   CreateTimeSpan 10, 20, 30
   CreateTimeSpan -10, 20, 30
   CreateTimeSpan 0, 0, 37230
   CreateTimeSpan 1000, 2000, 3000
   CreateTimeSpan 1000, -2000, -3000
   CreateTimeSpan 999999, 999999, 999999
End Sub

Private Sub CreateTimeSpan(ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long)
   Dim elapsedTime As ITimeSpan
   Set elapsedTime = TimeSpan.Create(Hours, Minutes, Seconds)

   ' Format the constructor for display.
   Dim ctor As String
   ctor = "TimeSpan( " & Hours & ", " & Minutes & ", " & Seconds & " )"
   
   ' Display the constructor and its value.
   Debug.Print ctor, elapsedTime.ToString()
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
