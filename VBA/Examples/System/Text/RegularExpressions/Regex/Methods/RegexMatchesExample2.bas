Attribute VB_Name = "RegexMatchesExample2"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 3, 2024
'@LastModified February 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.matches?view=netframework-4.8.1#system-text-regularexpressions-regex-matches(system-string-system-int32)

Option Explicit

''
' The following example uses the Match(String) method to find the first word in
' a sentence that ends in "es", and then calls the Matches(String, Int32)
' method to identify any additional words that end in "es".
''
Public Sub RegexMatchesExample2()
    Dim pattern As String
    pattern = "\b\w+es\b"
    Dim rgx As DotNetLib.Regex
    Set rgx = Regex.Create(pattern)
    Dim sentence As String
    sentence = "Who writes these notes and uses our paper?"
    
    ' Get the first match.
    Dim pvtMatch As DotNetLib.Match
    Set pvtMatch = rgx.Match(sentence)
    If (pvtMatch.Success) Then
        Debug.Print VBString.Format("Found first 'es' in '{0}' at position {1}", _
                           pvtMatch.value, pvtMatch.index)
        ' Get any additional matches.
        Dim varMatch As Variant
        For Each varMatch In rgx.Matches2(sentence, pvtMatch.index + pvtMatch.Length)
            Dim m As DotNetLib.Match
            Set m = varMatch
            Debug.Print VBString.Format("Also found '{0}' at position {1}", _
                              m.value, m.index)
        Next
    End If
End Sub

' The example displays the following output:
'       Found first 'es' in 'writes' at position 4
'       Also found 'notes' at position 17
'       Also found 'uses' at position 27
