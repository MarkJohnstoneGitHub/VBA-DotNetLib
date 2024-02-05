Attribute VB_Name = "RegexReplace5Example"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 5, 2024
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.replace?view=netframework-4.8.1#system-text-regularexpressions-regex-replace(system-string-system-text-regularexpressions-matchevaluator-system-int32)

'@Dependencies ReverseLetterMatchEvaluator.cls

Option Explicit

''
' The following example uses a regular expression to deliberately misspell half
' of the words in a list. It uses the regular expression \w*(ie|ei)\w* to match
' words that include the characters "ie" or "ei". It passes the first half of
' the matching words to the ReverseLetter method, which, in turn, uses the
' Replace(String, String, String, RegexOptions) method to reverse "i" and "e"
' in the matched string. The remaining words remain unchanged.
''
Public Sub RegexReplace5Example()
    Dim pvtInput As DotNetLib.String
    Set pvtInput = Strings.Create("deceive relieve achieve belief fierce receive")
    Dim pattern As String
    pattern = "\w*(ie|ei)\w*"
    Dim rgx As DotNetLib.Regex
    Set rgx = Regex.Create(pattern, RegexOptions.RegexOptions_IgnoreCase)
    Debug.Print "Original string: " & pvtInput.ToString

    Dim pvtResult As String
    pvtResult = rgx.Replace5(pvtInput.ToString, MatchEvaluator.Create(ReverseLetterMatchEvaluator), _
                            (UBound(pvtInput.Split(" ")) + 1) / 2)
    Debug.Print "Returned string: " + pvtResult
End Sub

' The example displays the following output:
'    Original string: deceive relieve achieve belief fierce receive
'    Returned string: decieve releive acheive belief fierce receive
