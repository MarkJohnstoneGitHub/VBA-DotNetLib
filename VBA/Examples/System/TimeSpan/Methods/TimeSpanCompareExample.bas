Attribute VB_Name = "TimeSpanCompareExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 17, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.compare?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the Compare method to compare several TimeSpan
' objects with a TimeSpan object whose value is a 2-hour time interval.
''
Public Sub TimeSpanCompare()
    ' Define a time interval equal to two hours.
    Dim baseInterval As DotNetLib.TimeSpan
    Set baseInterval = TimeSpan.Create(2, 0, 0)
    
    ' Define an array of time intervals to compare with
    ' the base interval.
    Dim spans() As DotNetLib.TimeSpan
    ObjectArray.CreateInitialize1D spans, _
                TimeSpan.FromSeconds(-2.5), _
                TimeSpan.FromMinutes(20), _
                TimeSpan.FromHours(1), _
                TimeSpan.FromMinutes(90), _
                baseInterval, _
                TimeSpan.FromDays(0.5), _
                TimeSpan.FromDays(1)
                  
    ' Compare the time intervals.
    Dim varSpan As Variant
    For Each varSpan In spans
        Dim span As DotNetLib.TimeSpan
        Set span = varSpan
        Dim result As Long
        result = TimeSpan.Compare(baseInterval, span)
        Debug.Print VBString.Format("{0} {1} {2} (Compare returns {3})", _
                     baseInterval, _
                     IIf(result = 1, ">", IIf(result = 0, "=", "<")), _
                     span, result)
   Next
End Sub

' The example displays the following output:
'       02:00:00 > -00:00:02.5000000 (Compare returns 1)
'       02:00:00 > 00:20:00 (Compare returns 1)
'       02:00:00 > 01:00:00 (Compare returns 1)
'       02:00:00 > 01:30:00 (Compare returns 1)
'       02:00:00 = 02:00:00 (Compare returns 0)
'       02:00:00 < 12:00:00 (Compare returns -1)
'       02:00:00 < 1.00:00:00 (Compare returns -1)

