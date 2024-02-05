Attribute VB_Name = "RegexMatchesExample3"
'@IgnoreModule EmptyIfBlock
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 3, 2024
'@LastModified February 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.matches?view=netframework-4.8.1#system-text-regularexpressions-regex-matches(system-string-system-string-system-text-regularexpressions-regexoptions-system-timespan)

Option Explicit

''
' The following example calls the Matches(String, String, RegexOptions, TimeSpan)
' method to perform a case-sensitive comparison that matches any word in a
' sentence that ends in "es".
' It then calls the Matches(String, String, RegexOptions, TimeSpan) method to
' perform a case-insensitive comparison of the pattern with the input string.
' In both cases, the time-out interval is set to one second. As the output shows,
' the two methods return different results.
''
Public Sub RegexMatchesExample3()
    Dim pattern As String
    pattern = "\b\w+es\b"
    Dim sentence As String
    sentence = "NOTES: Any notes or comments are optional."
    
    ' Call Matches method without specifying any options.
    On Error Resume Next
    Dim varMatch As Variant
    Dim pvtMatch As DotNetLib.Match
    For Each varMatch In Regex.Matches2(sentence, pattern, _
                                        RegexOptions.RegexOptions_None, _
                                        TimeSpan.FromSeconds(1))
        If Err.Number = 0 Then
            Set pvtMatch = varMatch
            Debug.Print VBString.Format("Found '{0}' at position {1}", _
                                  pvtMatch.value, pvtMatch.index)
        ElseIf Err.Number = RegexMatchTimeoutException Then
            ' Do Nothing: Assume that timeout represents no match.
        End If
    Next
    On Error GoTo 0 'Stop code and display error
    Debug.Print

    ' Call Matches method for case-insensitive matching.
    On Error Resume Next
    For Each varMatch In Regex.Matches(sentence, pattern, RegexOptions.RegexOptions_IgnoreCase)
        If Err.Number = 0 Then
            Set pvtMatch = varMatch
            Debug.Print VBString.Format("Found '{0}' at position {1}", _
                                        pvtMatch.value, pvtMatch.index)
        ElseIf Err.Number = RegexMatchTimeoutException Then
            ' Do Nothing: Assume that timeout represents no match.
        End If
    Next
    On Error GoTo 0 'Stop code and display error
End Sub

' The example displays the following output:
'       Found 'notes' at position 11
'
'       Found 'NOTES' at position 0
'       Found 'notes' at position 11

