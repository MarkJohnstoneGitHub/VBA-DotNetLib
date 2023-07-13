Attribute VB_Name = "DateTimeMinValueExample"
'Rubberduck annotations
'@Folder "VBADotNetLib.Examples.DateTime.Properties"

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 09, 2023

Option Explicit

Public Sub DateTimeMinValueField()
    Dim date1 As DateTime
    Set date1 = New DateTime
    
    If (date1.Equals(DateTime.MinValue)) Then
        Debug.Print date1.ToString & " (Equals Date.MinValue)"
    End If
    ' The example displays the following output:
    '    1/1/0001 12:00:00 AM  (Equals Date.MinValue)
End Sub
