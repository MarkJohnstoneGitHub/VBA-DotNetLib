Attribute VB_Name = "TimeSpanEqualityExample"
'@Folder "Examples.System.TimeSpan.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.op_equality?view=netframework-4.8.1#examples

Option Explicit

Private Const dataFmt As String = "{0,34}    {1}"

''
' The following example compares several TimeSpan objects to a reference TimeSpan
' using the Equality operator.
''
Public Sub TimeSpanEquality()
    Dim pvtLeft As DotNetLib.TimeSpan
    Set pvtLeft = TimeSpan.Create(2, 0, 0)

    Debug.Print VBString.Unescape( _
        "This example of the TimeSpan relational operators " + _
        "generates \nthe following output. It creates several " + _
        "different TimeSpan \nobjects and compares them with " + _
        "a 2-hour TimeSpan.\n")
    Debug.Print VBString.Format(dataFmt, _
        "Left: TimeSpan( 2, 0, 0 )", pvtLeft)

    ' Create objects to compare with a 2-hour TimeSpan.
    CompareTimeSpans pvtLeft, TimeSpan.Create(0, 120, 0), "TimeSpan( 0, 120, 0 )"
    CompareTimeSpans pvtLeft, TimeSpan.Create(2, 0, 1), "TimeSpan( 2, 0, 1 )"
    CompareTimeSpans pvtLeft, TimeSpan.Create(2, 0, -1), "TimeSpan( 2, 0, -1 )"
    CompareTimeSpans pvtLeft, TimeSpan.FromDays(1 / 12), "TimeSpan.FromDays( 1 / 12 )"
End Sub

Private Sub CompareTimeSpans(ByVal pLeft As DotNetLib.TimeSpan, ByVal pRight As DotNetLib.TimeSpan, ByVal rightText As String)
    Debug.Print
    Debug.Print VBString.Format(dataFmt, "Right: " + rightText, pRight)
    Debug.Print VBString.Format(dataFmt, "Left == Right", TimeSpan.Equality(pLeft, pRight))
    Debug.Print VBString.Format(dataFmt, "Left >  Right", TimeSpan.GreaterThan(pLeft, pRight))
    Debug.Print VBString.Format(dataFmt, "Left >= Right", TimeSpan.GreaterThanOrEqual(pLeft, pRight))
    Debug.Print VBString.Format(dataFmt, "Left != Right", TimeSpan.Inequality(pLeft, pRight))
    Debug.Print VBString.Format(dataFmt, "Left <  Right", TimeSpan.LessThan(pLeft, pRight))
    Debug.Print VBString.Format(dataFmt, "Left <= Right", TimeSpan.LessThanOrEqual(pLeft, pRight))
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

