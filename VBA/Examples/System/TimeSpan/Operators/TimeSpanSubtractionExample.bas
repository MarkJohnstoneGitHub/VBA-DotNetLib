Attribute VB_Name = "TimeSpanSubtractionExample"
'@Folder "Examples.System.TimeSpan.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.op_subtraction?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the TimeSpan subtraction operator to calculate the
' total length of the weekly work day. It also uses the TimeSpan addition
' operator to compute the total time of the daily breaks before using it in a
' subtraction operation to compute the total actual daily working time.
''
Public Sub TimeSpanSubtraction()
    Dim startWork As DotNetLib.TimeSpan
    Set startWork = TimeSpan.Create(8, 0, 0)
    Dim endWork As DotNetLib.TimeSpan
    Set endWork = TimeSpan.Create(18, 30, 0)
    Dim lunchBreak As DotNetLib.TimeSpan
    Set lunchBreak = TimeSpan.Create(1, 0, 0)
    Dim breaks As DotNetLib.TimeSpan
    Set breaks = TimeSpan.Create(0, 30, 0)
   
    Debug.Print VBString.Format("Length of work day: {0}", _
                      TimeSpan.Subtraction(endWork, startWork))
    Debug.Print VBString.Format("Actual time worked: {0}", _
                      TimeSpan.Subtraction(TimeSpan.Subtraction(endWork, startWork), TimeSpan.Addition(lunchBreak, breaks)))
End Sub

' The example displays the following output:
'     Length of work day: 10:30:00
'     Actual time worked: 09:00:00

