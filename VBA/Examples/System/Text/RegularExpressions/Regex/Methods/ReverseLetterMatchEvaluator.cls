VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReverseLetterMatchEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 4, 2024
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.replace?view=netframework-4.8.1#system-text-regularexpressions-regex-replace(system-string-system-text-regularexpressions-matchevaluator-system-int32)

Option Explicit

Implements DotNetLib.IMatchEvaluator

'----------------------------------------------------------------------
' Interface IMatchEvaluator
'----------------------------------------------------------------------

''
' ReverseLetter method, which uses the Replace(String, String, String, RegexOptions)
' method to reverse "i" and "e" in the matched string.
' @Remarks Used for the RegexReplace5Example.bas
''
Private Function IMatchEvaluator_MatchEvaluator(ByVal pMatch As DotNetLib.Match) As String
    IMatchEvaluator_MatchEvaluator = Regex.Replace(pMatch.value, "([ie])([ie])", "$2$1", _
                                        RegexOptions.RegexOptions_IgnoreCase)
End Function

