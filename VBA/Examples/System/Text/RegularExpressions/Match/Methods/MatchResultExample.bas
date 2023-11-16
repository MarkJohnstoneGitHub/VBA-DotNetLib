Attribute VB_Name = "MatchResultExample"
'@Folder("Examples.System.Text.RegularExpressions.Match.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 1, 2023
'@LastModified October 1, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match.result?view=netframework-4.8.1#examples

Option Explicit

''
' The following example replaces the hyphens that begin and end a parenthetical
' expression with parentheses.
''
Public Sub MatchResult()
    Dim pattern As String
    pattern = "--(.+?)--"
    Dim replacement As String
    replacement = "($1)"
    Dim strInput As String
    strInput = "He said--decisively--that the time--whatever time it was--had come."
    
    Dim varMatch As Variant
    For Each varMatch In Regex.Matches(strInput, pattern)
        Dim pvtMatch As DotNetLib.Match
        Set pvtMatch = varMatch
        Dim pvtResult As String
        pvtResult = pvtMatch.result(replacement)
        Debug.Print pvtResult
    Next
End Sub

' The example displays the following output:
'       (decisively)
'       (whatever time it was)
