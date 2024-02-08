Attribute VB_Name = "MatchEvaluatorExample"
'@Folder("Examples.System.Text.RegularExpressions.MatchEvaluator")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 8, 2024
'@LastModified February 8, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchevaluator?view=netframework-4.8.1#examples

'Dependencies ReplaceCCMatchEvaluator.cls

Option Explicit

''
' The following code example uses the MatchEvaluator delegate to replace every
' matched group of characters with the number of the match occurrence.
''
Public Sub MatchEvaluatorExample()
    Dim sInput As String
    Dim sRegex As String

    ' The string to search.
    sInput = "aabbccddeeffcccgghhcccciijjcccckkcc"
    
    ' A very simple regular expression.
    sRegex = "cc"

    Dim r As DotNetLib.Regex
    Set r = Regex.Create(sRegex)
    
    ' Assign the replace method to the MatchEvaluator delegate.
    Dim myEvaluator As DotNetLib.MatchEvaluator
    Set myEvaluator = MatchEvaluator.Create(New ReplaceCCMatchEvaluator)

    ' Write out the original string.
    Debug.Print sInput
    
    ' Replace matched characters using the delegate method.
    sInput = r.Replace4(sInput, myEvaluator)
    
    ' Write out the modified string.
    Debug.Print sInput
End Sub

' The example displays the following output:
'       aabbccddeeffcccgghhcccciijjcccckkcc
'       aabb11ddeeff22cgghh3344iijj5566kk77
