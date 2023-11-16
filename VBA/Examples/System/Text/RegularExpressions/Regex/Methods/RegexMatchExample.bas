Attribute VB_Name = "RegexMatchExample"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 29, 2023
'@LastModified September 29, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.match?view=netframework-4.8.1#system-text-regularexpressions-regex-match(system-string-system-string-system-text-regularexpressions-regexoptions)

Option Explicit

''
' The following example defines a regular expression that matches words
' beginning with the letter "a". It uses the RegexOptions.IgnoreCase option to
' ensure that the regular expression locates words beginning with both an
' uppercase "a" and a lowercase "a".
''
Public Sub RegexMatch()
    Dim pattern As String
    pattern = "\ba\w*\b"
    
    Dim strInput As String
    strInput = "An extraordinary day dawns with each new day."
    
    Dim m As DotNetLib.Match
    Set m = Regex.Match(strInput, pattern, RegexOptions.RegexOptions_IgnoreCase)
    If (m.Success) Then
        Debug.Print VBAString.Format("Found '{0}' at position {1}.", m.value, m.index)
    End If
End Sub

' The example displays the following output:
'        Found 'An' at position 0.

