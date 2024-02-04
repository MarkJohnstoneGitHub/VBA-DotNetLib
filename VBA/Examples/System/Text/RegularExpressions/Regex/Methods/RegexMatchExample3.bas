Attribute VB_Name = "RegexMatchExample3"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 3, 2024
'@LastModified February 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.match?view=netframework-4.8.1#system-text-regularexpressions-regex-match(system-string-system-string)

Option Explicit

''
' The following example calls the Match(String, String) method to find the
' first word that contains at least one z character, and then calls the
' Match.NextMatch method to find any additional matches.
''
Public Sub RegexMatchExample3()
    Dim pvtInput  As String
    pvtInput = "ablaze beagle choral dozen elementary fanatic " + _
                "glaze hunger inept jazz kitchen lemon minus " + _
                "night optical pizza quiz restoration stamina " + _
                "train unrest vertical whiz xray yellow zealous"
    Dim pattern As String
    pattern = "\b\w*z+\w*\b"
    Dim m As DotNetLib.Match
    Set m = Regex.Match(pvtInput, pattern)
    Do While (m.Success)
        Debug.Print VBString.Format("'{0}' found at position {1}", m.value, m.index)
        Set m = m.NextMatch()
    Loop
End Sub

' The example displays the following output:
'    'ablaze' found at position 0
'    'dozen' found at position 21
'    'glaze' found at position 46
'    'jazz' found at position 65
'    'pizza' found at position 104
'    'quiz' found at position 110
'    'whiz' found at position 157
'    'zealous' found at position 174
