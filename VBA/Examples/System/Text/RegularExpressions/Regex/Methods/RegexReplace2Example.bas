Attribute VB_Name = "RegexReplace2Example"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 5, 2024
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.replace?view=netframework-4.8.1#system-text-regularexpressions-regex-replace(system-string-system-string-system-int32)

Option Explicit

''
' The following example replaces the first five occurrences of duplicated
' characters with a single character. The regular expression pattern
' (\w)\1 matches consecutive occurrences of a single character and assigns
' the first occurrence to the first capturing group. The replacement pattern $1
' replaces the entire match with the first captured group.
''
Public Sub RegexReplace2Example()
    Dim str As String
    str = "aabccdeefgghiijkklmm"
    Dim pattern As String
    pattern = "(\w)\1"
    Dim replacement As String
    replacement = "$1"
    Dim rgx As DotNetLib.Regex
    Set rgx = Regex.Create(pattern)
    
    Dim pvtResult As String
    pvtResult = rgx.Replace2(str, replacement, 5)
    Debug.Print VBString.Format("Original String:    '{0}'", str)
    Debug.Print VBString.Format("Replacement String: '{0}'", pvtResult)
End Sub

' The example displays the following output:
'       Original String:    'aabccdeefgghiijkklmm'
'       Replacement String: 'abcdefghijkklmm'
