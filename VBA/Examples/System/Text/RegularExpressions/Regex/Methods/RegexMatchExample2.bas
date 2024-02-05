Attribute VB_Name = "RegexMatchExample2"
'@IgnoreModule IndexedDefaultMemberAccess
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 3, 2024
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.match?view=netframework-4.8.1#system-text-regularexpressions-regex-match(system-string)

Option Explicit

''
' The following example finds regular expression pattern matches in a string,
' then lists the matched groups, captures, and capture positions.
''
Public Sub RegexMatchExample2()
    Dim text As String
    text = "One car red car blue car"
    Dim pat As String
    pat = "(\w+)\s+(car)"
    
    ' Instantiate the regular expression object.
    Dim r As DotNetLib.Regex
    Set r = Regex.Create(pat, RegexOptions.RegexOptions_IgnoreCase)
    
    ' Match the regular expression pattern against a text string.
    Dim m As DotNetLib.Match
    Set m = r.Match(text)
    Dim matchCount As Long
    matchCount = 0
    Do While (m.Success)
        matchCount = matchCount + 1
        Debug.Print "Match" & matchCount
        Dim i As Long
        For i = 1 To 2
            Dim g As DotNetLib.Group
            Set g = m.Groups(i)
            Debug.Print "Group" & i&; "='" & g.ToString & "'"
            Dim cc As DotNetLib.CaptureCollection
            Set cc = g.Captures
            Dim j As Long
            For j = 0 To cc.Count - 1
                Dim c As DotNetLib.Capture
                Set c = cc(j)
                Debug.Print "Capture" & j&; "='" & c.ToString & "', Position=" & c.index
            Next
        Next
        Set m = m.NextMatch()
    Loop
End Sub

' This example displays the following output:
'       Match1
'       Group1='One'
'       Capture0='One', Position=0
'       Group2='car'
'       Capture0='car', Position=4
'       Match2
'       Group1='red'
'       Capture0='red', Position=8
'       Group2='car'
'       Capture0='car', Position=12
'       Match3
'       Group1='blue'
'       Capture0='blue', Position=16
'       Group2='car'
'       Capture0='car', Position=21
