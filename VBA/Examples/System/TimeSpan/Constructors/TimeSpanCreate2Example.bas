Attribute VB_Name = "TimeSpanCreate2Example"
'@Folder "Examples.System.TimeSpan.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 15, 2023
'@LastModified September 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.-ctor?view=netframework-4.8.1#system-timespan-ctor(system-int32-system-int32-system-int32-system-int32)

Option Explicit

'@Description("The following example creates several TimeSpan objects using the constructor overload that initializes a TimeSpan to a specified number of days, hours, minutes, and seconds.")
Public Sub TimeSpanCreate2()
Attribute TimeSpanCreate2.VB_Description = "The following example creates several TimeSpan objects using the constructor overload that initializes a TimeSpan to a specified number of days, hours, minutes, and seconds."
    Debug.Print VBAString.Format("{0,-44}{1,16}", "Constructor", "Value")
    Debug.Print VBAString.Format("{0,-44}{1,16}", "-----------", "-----")
    CreateTimeSpan 10, 20, 30, 40
    CreateTimeSpan -10, 20, 30, 40
    CreateTimeSpan 0, 0, 0, 937840
    CreateTimeSpan 1000, 2000, 3000, 4000
    CreateTimeSpan 1000, -2000, -3000, -4000
    CreateTimeSpan 999999, 999999, 999999, 999999
End Sub

' Create a TimeSpan object and display its value.
Private Sub CreateTimeSpan(ByVal pDays As Long, ByVal pHours As Long, ByVal pMinutes As Long, ByVal pSeconds As Long)
    Dim elapsedTime As ITimeSpan
    Set elapsedTime = TimeSpan.Create2(pDays, pHours, pMinutes, pSeconds)

    ' Format the constructor for display.
    Dim ctor As String
    ctor = VBAString.Format("TimeSpan( {0}, {1}, {2}, {3} )", pDays, pHours, pMinutes, pSeconds)

    ' Display the constructor and its value.
    Debug.Print VBAString.Format("{0,-44}{1,16}", ctor, elapsedTime.ToString())
End Sub

' The example displays the following output:
'       Constructor                                            Value
'       -----------                                            -----
'       TimeSpan( 10, 20, 30, 40 )                       10.20:30:40
'       TimeSpan( -10, 20, 30, 40 )                      -9.03:29:20
'       TimeSpan( 0, 0, 0, 937840 )                      10.20:30:40
'       TimeSpan( 1000, 2000, 3000, 4000 )             1085.11:06:40
'       TimeSpan( 1000, -2000, -3000, -4000 )           914.12:53:20
'       TimeSpan( 999999, 999999, 999999, 999999 )  1042371.15:25:39
