Attribute VB_Name = "RegexReplace4Example"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 4, 2024
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.replace?view=netframework-4.8.1#system-text-regularexpressions-regex-replace(system-string-system-text-regularexpressions-matchevaluator)

'@Dependencies CapTextMatchEvaluator.cls

Option Explicit

''
' The following code example displays an original string, matches each word in
' the original string, converts the first character of each match to uppercase,
' then displays the converted string.
''
Public Sub RegexReplace4Example()
    Dim text As String
    text = "four score and seven years ago"

    Debug.Print "text=[{text}]"
    Debug.Print VBString.Format("text=[{0}]", text)

    Dim rx As DotNetLib.Regex
    Set rx = Regex.Create("\w+")

    Dim pvtResult As String

    pvtResult = rx.Replace4(text, MatchEvaluator.Create(CapTextMatchEvaluator))
    Debug.Print VBString.Format("result=[{0}]", pvtResult)
End Sub

' The example displays the following output:
'       text=[four score and seven years ago]
'       result=[Four Score And Seven Years Ago]
