Attribute VB_Name = "MatchGroupsExample"
'@Folder("Examples.System.Text.RegularExpressions.Match.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 1, 2023
'@LastModified October 1, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match.groups?view=netframework-4.8.1#examples

Option Explicit

''
' The following example attempts to match a regular expression pattern against
' a sample string. The example uses the Groups property to store information
' that is retrieved by the match for display to the console.
''
Public Sub MatchGroups()
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
    
    While (m.Success)
        matchCount = matchCount + 1
        Debug.Print Strings.Format("Match" & (matchCount))
        Dim i As Long
        For i = 1 To 2
            Dim g As DotNetLib.Group
            Set g = m.Groups(i)
            Debug.Print Strings.Format("Group" & i & "='" & g & "'")
            Dim cc  As DotNetLib.CaptureCollection
            Set cc = g.Captures
            Dim j As Long
            For j = 0 To cc.Count - 1
                Dim c As DotNetLib.Capture
                Set c = cc(j)
                Debug.Print Strings.Format("Capture" & j & "='" & c & "', Position=" & c.Index)
            Next j
        Next i
         Set m = m.NextMatch()
    Wend
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
