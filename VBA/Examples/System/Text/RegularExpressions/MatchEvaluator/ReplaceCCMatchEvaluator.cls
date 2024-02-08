VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReplaceCCMatchEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Examples.System.Text.RegularExpressions.MatchEvaluator")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 8, 2024
'@LastModified February 8, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchevaluator?view=netframework-4.8.1#examples

Option Explicit

Implements DotNetLib.IMatchEvaluator

'----------------------------------------------------------------------
' Interface IMatchEvaluator
'----------------------------------------------------------------------

''
' The following code example uses the MatchEvaluator delegate to replace every
' matched group of characters with the number of the match occurrence.
'
' @Remarks Used for the MatchEvaluatorExample.bas
''
Private Function IMatchEvaluator_MatchEvaluator(ByVal pMatch As DotNetLib.Match) As String
    Static i As Long
    i = i + 1
    IMatchEvaluator_MatchEvaluator = CStr(i) & CStr(i)
End Function
