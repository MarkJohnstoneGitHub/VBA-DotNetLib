Attribute VB_Name = "TimeSpanCompareTo2Example"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 4, 2023
'@LastModified January 17, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.compareto?view=netframework-4.8.1#system-timespan-compareto(system-object)

Option Explicit

''
' Compares the value of the current DateTimeOffset object with another object
' of the same type.
''
Public Sub TimeSpanCompareTo2()
    Dim pvtLeft As DotNetLib.TimeSpan
    Set pvtLeft = TimeSpan.Create(0, 5, 0)
    Debug.Print VBString.Unescape( _
        "This example of the TimeSpan.Equals( Object ) " + _
        "and \nTimeSpan.CompareTo( Object ) methods generates " + _
        "the \nfollowing output by creating several different " + _
        "TimeSpan \nobjects and comparing them with a " + _
        "5-minute TimeSpan.\n")
    Debug.Print VBString.Format(VBString.Unescape("{0,-33}{1}\n"), _
            "Left: TimeSpan( 0, 5, 0 )", pvtLeft)
    
    ' Create objects to compare with a 5-minute TimeSpan.
    Call CompTimeSpanToObject(pvtLeft, TimeSpan.Create(0, 0, 300), "TimeSpan( 0, 0, 300 )")
    Call CompTimeSpanToObject(pvtLeft, TimeSpan.Create(0, 5, 1), "TimeSpan( 0, 5, 1 )")
    Call CompTimeSpanToObject(pvtLeft, TimeSpan.Create(0, 5, -1), "TimeSpan( 0, 5, -1 )")
    Call CompTimeSpanToObject(pvtLeft, TimeSpan.CreateFromTicks(3000000000#), "TimeSpan( 3000000000 )")
    Call CompTimeSpanToObject(pvtLeft, DateTime.Now, "DateTime.Now")
End Sub

' Compare the TimeSpan to the Object parameters,
' and display the Object parameters with the results.
Private Sub CompTimeSpanToObject(ByVal pLeft As DotNetLib.TimeSpan, ByVal pRight As Object, ByVal rightText As String)
    Debug.Print VBString.Format("{0,-33}{1}", "Object: " + rightText, _
                                pRight)
    Debug.Print VBString.Format("{0,-33}{1}", "Left.Equals( Object )", _
                                pLeft.Equals2(pRight))
    Debug.Print VBString.Format("{0,-33}", "Left.CompareTo( Object )");
    On Error Resume Next
    Debug.Print pLeft.CompareTo2(pRight) & VBA.vbNewLine
    If Catch() Then
        Debug.Print "Error: "; Err.Description & VBA.vbNewLine
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

'/*
'This example of the TimeSpan.Equals( Object ) and
'TimeSpan.CompareTo( Object ) methods generates the
'following output by creating several different TimeSpan
'objects and comparing them with a 5-minute TimeSpan.
'
'Left: TimeSpan( 0, 5, 0 )        00:05:00
'
'Object: TimeSpan( 0, 0, 300 )    00:05:00
'Left.Equals( Object )            True
'Left.CompareTo( Object )         0
'
'Object: TimeSpan( 0, 5, 1 )      00:05:01
'Left.Equals( Object )            False
'Left.CompareTo (Object) - 1
'
'Object: TimeSpan( 0, 5, -1 )     00:04:59
'Left.Equals( Object )            False
'Left.CompareTo( Object )         1
'
'Object: TimeSpan( 3000000000 )   00:05:00
'Left.Equals( Object )            True
'Left.CompareTo( Object )         0
'
'Object: long 3000000000L         3000000000
'Left.Equals( Object )            False
'Left.CompareTo( Object )         Error: Object must be of type TimeSpan.
'
'Object: string "00:05:00"        00:05:00
'Left.Equals( Object )            False
'Left.CompareTo( Object )         Error: Object must be of type TimeSpan.
'*/

