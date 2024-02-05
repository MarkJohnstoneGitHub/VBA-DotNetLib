Attribute VB_Name = "RegexReplaceStaticExample"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 5, 2024
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.replace?view=netframework-4.8.1#system-text-regularexpressions-regex-replace(system-string-system-string)

Option Explicit

''
' The following example defines a regular expression, \s+, that matches one or
' more white-space characters. The replacement string, " ", replaces them with
' a single space character.
''
Public Sub RegexReplaceStaticExample()
    Dim pvtInput As String
    pvtInput = "This is   text with   far  too   much   " + _
                "white space."
    Dim pattern As String
    pattern = VBString.Unescape("\\s+")
    Dim replacement As String
    replacement = " "
    Dim rgx As DotNetLib.Regex
    Set rgx = Regex.Create(pattern)
    Dim pvtResult As String
    pvtResult = rgx.Replace(pvtInput, replacement)

    Debug.Print VBString.Format("Original String: {0}", pvtInput)
        Debug.Print VBString.Format("Replacement String: {0}", pvtResult)
End Sub

' The example displays the following output:
'       Original String: This is   text with   far  too   much   white space.
'       Replacement String: This is text with far too much white space.
