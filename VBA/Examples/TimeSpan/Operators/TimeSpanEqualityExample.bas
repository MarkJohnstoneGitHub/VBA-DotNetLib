Attribute VB_Name = "TimeSpanEqualityExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Operators")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.op_equality?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example compares several TimeSpan objects to a reference TimeSpan using the Equality operator.")
Public Sub TimeSpanEquality()
Attribute TimeSpanEquality.VB_Description = "The following example compares several TimeSpan objects to a reference TimeSpan using the Equality operator."
   Dim left As ITimeSpan
   Set left = TimeSpan.Create(2, 0, 0)
   
   Debug.Print "This example of the TimeSpan relational operators generates " & VBA.vbNewLine & _
               "the following output. It creates several different TimeSpan " & VBA.vbNewLine & _
               "objects and compares them with " & _
               "a 2-hour TimeSpan." & VBA.vbNewLine
   Debug.Print "Left: TimeSpan( 2, 0, 0 )" & "     " & left.ToString()

   ' Create objects to compare with a 2-hour TimeSpan.
   CompareTimeSpans left, TimeSpan.Create(0, 120, 0), "TimeSpan( 0, 120, 0 )"
   CompareTimeSpans left, TimeSpan.Create(2, 0, 1), "TimeSpan( 2, 0, 1 )"
   CompareTimeSpans left, TimeSpan.Create(2, 0, -1), "TimeSpan( 2, 0, -1 )"
   CompareTimeSpans left, TimeSpan.FromDays(1 / 12), "TimeSpan.FromDays( 1 / 12 )"
End Sub

Private Sub CompareTimeSpans(ByVal left As ITimeSpan, ByVal right As ITimeSpan, ByVal rightText As String)
   Debug.Print
   Debug.Print "Right: " + rightText & "     " & right.ToString()
   Debug.Print "Left == Right", TimeSpan.Equality(left, right)
   Debug.Print "Left >  Right", TimeSpan.GreaterThan(left, right)
   Debug.Print "Left >= Right", TimeSpan.GreaterThanOrEqual(left, right)
   Debug.Print "Left != Right", TimeSpan.Inequality(left, right)
   Debug.Print "Left <  Right", TimeSpan.LessThan(left, right)
   Debug.Print "Left <= Right", TimeSpan.LessThanOrEqual(left, right)
End Sub

'/*
'This example of the TimeSpan relational operators generates
'the following output. It creates several different TimeSpan
'objects and compares them with a 2-hour TimeSpan.
'
'         Left: TimeSpan( 2, 0, 0 )    02:00:00
'
'      Right: TimeSpan( 0, 120, 0 )    02:00:00
'                     Left == Right    True
'                     Left >  Right    False
'                     Left >= Right    True
'                     Left != Right    False
'                     Left <  Right    False
'                     Left <= Right    True
'
'        Right: TimeSpan( 2, 0, 1 )    02:00:01
'                     Left == Right    False
'                     Left >  Right    False
'                     Left >= Right    False
'                     Left != Right    True
'                     Left <  Right    True
'                     Left <= Right    True
'
'       Right: TimeSpan( 2, 0, -1 )    01:59:59
'                     Left == Right    False
'                     Left >  Right    True
'                     Left >= Right    True
'                     Left != Right    True
'                     Left <  Right    False
'                     Left <= Right    False
'
'Right: TimeSpan.FromDays( 1 / 12 )    02:00:00
'                     Left == Right    True
'                     Left >  Right    False
'                     Left >= Right    True
'                     Left != Right    False
'                     Left <  Right    False
'                     Left <= Right    True
'*/
