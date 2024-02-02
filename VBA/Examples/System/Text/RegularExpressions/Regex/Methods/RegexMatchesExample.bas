Attribute VB_Name = "RegexMatchesExample"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 3, 2024
'@LastModified February 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.matches?view=netframework-4.8.1#system-text-regularexpressions-regex-matches(system-string-system-string-system-text-regularexpressions-regexoptions)

Option Explicit

''
' The following example calls the Matches(String, String) method to identify any
' word in a sentence that ends in "es", and then calls the
' Matches(String, String, RegexOptions) method to perform a case-insensitive
' comparison of the pattern with the input string. As the output shows, the two
' methods return different results.
''
Public Sub RegexMatchesExample()
    Dim pattern As String
    pattern = "\b\w+es\b"
    Dim sentence As String
    sentence = "NOTES: Any notes or comments are optional."
    
    ' Call Matches method without specifying any options.
    Dim varMatch As Variant
    For Each varMatch In Regex.Matches(sentence, pattern)
        Dim pvtMatch As DotNetLib.Match
        Set pvtMatch = varMatch
        Debug.Print VBString.Format("Found '{0}' at position {1}", _
                           pvtMatch.value, pvtMatch.index)
    Next
    Debug.Print
    
    ' Call Matches method for case-insensitive matching.
    For Each varMatch In Regex.Matches(sentence, pattern, RegexOptions.RegexOptions_IgnoreCase)
        Set pvtMatch = varMatch
        Debug.Print VBString.Format("Found '{0}' at position {1}", _
                           pvtMatch.value, pvtMatch.index)
    Next
End Sub

' The example displays the following output:
'       Found 'notes' at position 11
'
'       Found 'NOTES' at position 0
'       Found 'notes' at position 11
