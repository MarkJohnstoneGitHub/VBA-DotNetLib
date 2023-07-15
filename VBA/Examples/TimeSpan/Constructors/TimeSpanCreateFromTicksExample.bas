Attribute VB_Name = "TimeSpanCreateFromTicksExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 15, 2023
'@LastModified July 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.-ctor?view=net-7.0#system-timespan-ctor(system-int64)

Option Explicit

Public Sub TimeSpanCreateFromTicks()
   Debug.Print "Constructor", "Value"
   CreateTimeSpan 1
   CreateTimeSpan 999999
   CreateTimeSpan -1000000000000#
   CreateTimeSpan 18012202000000#
   CreateTimeSpan "999999999999999999"
   CreateTimeSpan "1000000000000000000"
   
'This example of the TimeSpan( long ) constructor
'generates the following output.
'
'Constructor value
'TimeSpan( 1 )                            00:00:00.0000001
'TimeSpan( 999999 )                       00:00:00.0999999
'TimeSpan( -1000000000000 )            -1.03:46:40
'TimeSpan( 18012202000000 )            20.20:20:20.2000000
'TimeSpan( 999999999999999999 )   1157407.09:46:39.9999999
'TimeSpan( 1000000000000000000 )  1157407.09:46:40
End Sub

'@Description("'The following example creates several TimeSpan objects using the constructor overload that initializes a TimeSpan to a specified number of ticks.")
Private Sub CreateTimeSpan(ByVal pTicks As LongLong)
Attribute CreateTimeSpan.VB_Description = "'The following example creates several TimeSpan objects using the constructor overload that initializes a TimeSpan to a specified number of ticks."
   Dim elapsedTime As TimeSpan
   Set elapsedTime = TimeSpan.CreateFromTicks(pTicks)
   
   Dim ctor As String
   ctor = "TimeSpan( " & pTicks & " )"
   
   Dim elapsedStr As String
   elapsedStr = elapsedTime.ToString()
   
   Debug.Print ctor, elapsedStr
End Sub
