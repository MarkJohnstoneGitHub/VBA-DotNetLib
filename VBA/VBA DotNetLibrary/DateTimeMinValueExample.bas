Attribute VB_Name = "DateTimeMinValueExample"
'@Folder("Examples.DateTime")
Option Explicit

'@TODO Issue with equals, unitialized DateTime should be equal to MinValue?
Public Sub DateTimeMinValueField()
    Dim date1 As DateTime
    Set date1 = New DateTime
    
    If (date1.Equals(DateTime.MinValue)) Then
        Debug.Print date1.ToString & " (Equals Date.MinValue)"
    End If
    ' The example displays the following output:
    '    1/1/0001 12:00:00 AM  (Equals Date.MinValue)
End Sub
