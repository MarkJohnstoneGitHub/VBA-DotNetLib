Attribute VB_Name = "TimeSpanCreate2Example"
'@Folder("VBADotNetLib.Examples.TimeSpan.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 15, 2023
'@LastModified July 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.-ctor?view=net-7.0#system-timespan-ctor(system-int32-system-int32-system-int32-system-int32)

Option Explicit

Public Sub TimeSpanCreate()
   Debug.Print "Constructor", "Value"
   Debug.Print "-----------", "-----"
   CreateTimeSpan 10, 20, 30, 40
   CreateTimeSpan -10, 20, 30, 40
   CreateTimeSpan 0, 0, 0, 937840
   CreateTimeSpan 1000, 2000, 3000, 4000
   CreateTimeSpan 1000, -2000, -3000, -4000
   CreateTimeSpan 999999, 999999, 999999, 999999

' The example displays the following output:
'       Constructor                                            Value
'       -----------                                            -----
'       TimeSpan( 10, 20, 30, 40 )                       10.20:30:40
'       TimeSpan( -10, 20, 30, 40 )                      -9.03:29:20
'       TimeSpan( 0, 0, 0, 937840 )                      10.20:30:40
'       TimeSpan( 1000, 2000, 3000, 4000 )             1085.11:06:40
'       TimeSpan( 1000, -2000, -3000, -4000 )           914.12:53:20
'       TimeSpan( 999999, 999999, 999999, 999999 )  1042371.15:25:39
End Sub

' Create a TimeSpan object and display its value.
Private Sub CreateTimeSpan(ByVal Days As Long, ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long)
   Dim elapsedTime As TimeSpan
   Set elapsedTime = TimeSpan.Create2(Days, Hours, Minutes, Seconds)

   ' Format the constructor for display.
   Dim ctor As String
   ctor = "TimeSpan( " & Days & ", " & Hours & ", " & Minutes & ", " & Seconds & " )"
   
   ' Display the constructor and its value.
   Debug.Print ctor, elapsedTime.ToString()
End Sub

'The following example creates several TimeSpan objects using the constructor overload that initializes a TimeSpan to a specified number of days, hours, minutes, and seconds.

'Class Example
'{
'    // Create a TimeSpan object and display its value.
'    static void CreateTimeSpan( int days, int hours,
'        int minutes, int seconds )
'    {
'        TimeSpan elapsedTime =
'            new TimeSpan( days, hours, minutes, seconds );
'
'        // Format the constructor for display.
'        string ctor =
'            String.Format( "TimeSpan( {0}, {1}, {2}, {3} )",
'                days, hours, minutes, seconds);
'
'        // Display the constructor and its value.
'        Console.WriteLine( "{0,-44}{1,16}",
'            ctor, elapsedTime.ToString( ) );
'    }
'
'    static void Main( )
'    {
'        Console.WriteLine( "{0,-44}{1,16}", "Constructor", "Value" );
'        Console.WriteLine( "{0,-44}{1,16}", "-----------", "-----" );
'
'        CreateTimeSpan( 10, 20, 30, 40 );
'        CreateTimeSpan( -10, 20, 30, 40 );
'        CreateTimeSpan( 0, 0, 0, 937840 );
'        CreateTimeSpan( 1000, 2000, 3000, 4000 );
'        CreateTimeSpan( 1000, -2000, -3000, -4000 );
'        CreateTimeSpan( 999999, 999999, 999999, 999999 );
'    }
'}
'// The example displays the following output:
'//       Constructor                                            Value
'//       -----------                                            -----
'//       TimeSpan( 10, 20, 30, 40 )                       10.20:30:40
'//       TimeSpan( -10, 20, 30, 40 )                      -9.03:29:20
'//       TimeSpan( 0, 0, 0, 937840 )                      10.20:30:40
'//       TimeSpan( 1000, 2000, 3000, 4000 )             1085.11:06:40
'//       TimeSpan( 1000, -2000, -3000, -4000 )           914.12:53:20
'//       TimeSpan( 999999, 999999, 999999, 999999 )  1042371.15:25:39
