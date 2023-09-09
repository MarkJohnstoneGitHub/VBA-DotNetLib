Attribute VB_Name = "TimeSpanCompareTo2Example"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 4, 2023
'@LastModified September 9, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.compareto?view=netframework-4.8.1#system-timespan-compareto(system-object)

Option Explicit

'@Description("Compares the value of the current DateTimeOffset object with another object of the same type.")
Public Sub TimeSpanCompareTo2()
Attribute TimeSpanCompareTo2.VB_Description = "Compares the value of the current DateTimeOffset object with another object of the same type."
    Dim pvtLeft As ITimeSpan
    Set pvtLeft = TimeSpan.Create(0, 5, 0)
    Debug.Print "Left: TimeSpan( 0, 5, 0 )" & pvtLeft.ToString
    
    ' Create objects to compare with a 5-minute TimeSpan.
    CompTimeSpanToObject pvtLeft, TimeSpan.Create(0, 0, 300), "TimeSpan( 0, 0, 300 )"
    CompTimeSpanToObject pvtLeft, TimeSpan.Create(0, 5, 1), "TimeSpan( 0, 5, 1 )"
    CompTimeSpanToObject pvtLeft, TimeSpan.Create(0, 5, -1), "TimeSpan( 0, 5, -1 )"
    CompTimeSpanToObject pvtLeft, TimeSpan.CreateFromTicks(3000000000#), "TimeSpan( 3000000000 )"
    CompTimeSpanToObject pvtLeft, DateTime.Now, "DateTime.Now"
End Sub

' Compare the TimeSpan to the Object parameters,
' and display the Object parameters with the results.
Private Sub CompTimeSpanToObject(ByVal pLeft As ITimeSpan, ByVal pRight As Object, ByVal rightText As String)
    Debug.Print "Object: " & rightText
    Debug.Print "Left.Equals( Object ) :" & pLeft.Equals2(pRight)
    Debug.Print "Left.CompareTo( Object ) :";
    
    On Error Resume Next
    Debug.Print pLeft.CompareTo2(pRight) & VBA.vbNewLine
    If Catch() Then
        Debug.Print "Error: "; Err.Description & VBA.vbNewLine
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

'   Left: TimeSpan( 0, 5, 0 )00:05:00
'   Object: TimeSpan( 0, 0, 300 )
'   Left.Equals( Object ) :True
'   Left.CompareTo( Object ) :0
'
'   Object: TimeSpan( 0, 5, 1 )
'   Left.Equals( Object ) :False
'   Left.CompareTo( Object ) :-1
'
'   Object: TimeSpan( 0, 5, -1 )
'   Left.Equals( Object ) :False
'   Left.CompareTo( Object ) :1
'
'   Object: TimeSpan (3000000000)
'   Left.Equals( Object ) :True
'   Left.CompareTo( Object ) :0
'
'   Object:     DateTime.Now
'   Left.Equals( Object ) :False
'   Left.CompareTo( Object ) :Error: Object must be of type TimeSpan.
