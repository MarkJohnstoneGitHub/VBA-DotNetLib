Attribute VB_Name = "TimeSpanAdditionExample"
'@Folder "Examples.System.TimeSpan.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.op_addition?view=netframework-4.8.1

Option Explicit

''
' The Addition method defines the addition operator for TimeSpan values.
''
Public Sub TimeSpanAddition()
    Dim time1 As DotNetLib.TimeSpan
    Set time1 = TimeSpan.Create2(1, 0, 0, 0)     ' TimeSpan equivalent to 1 day.
    Dim time2 As DotNetLib.TimeSpan
    Set time2 = TimeSpan.Create(12, 0, 0)        ' TimeSpan equivalent to 1/2 day.
    Dim time3 As DotNetLib.TimeSpan
    Set time3 = TimeSpan.Addition(time1, time2)  ' Add the two time spans.
   
    Debug.Print VBString.Format(VBString.Unescape("  {0,12}\n +  {1,10}\n   {3}\n    {2,10}"), _
                  time1, time2, time3, Strings.Create2("_", 10))
End Sub

' The example displays the following output:
'           1.00:00:00
'        +    12:00:00
'          __________
'           1.12:00:00

