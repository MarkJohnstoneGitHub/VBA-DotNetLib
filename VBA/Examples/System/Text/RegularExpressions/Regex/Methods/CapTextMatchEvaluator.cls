VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CapTextMatchEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 4, 2024
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.replace?view=netframework-4.8.1#system-text-regularexpressions-regex-replace(system-string-system-text-regularexpressions-matchevaluator)

Option Explicit

Implements DotNetLib.IMatchEvaluator

'----------------------------------------------------------------------
' Interface IMatchEvaluator
'----------------------------------------------------------------------

''
' Converts the first character of each match to uppercase and returns the
' converted string.
'
' @Remarks Used for the RegexReplace4Example.bas
''
Private Function IMatchEvaluator_MatchEvaluator(ByVal pMatch As DotNetLib.Match) As String
    Dim pvtWord As String
    pvtWord = pMatch.ToString
    ' If the first char is lower case...
    Dim firstChar As String
    firstChar = Left$(pvtWord, 1)
    If firstChar <> UCase$(firstChar) Then
        pvtWord = StrConv(pvtWord, VBA.vbProperCase)
    End If
    IMatchEvaluator_MatchEvaluator = pvtWord
End Function

