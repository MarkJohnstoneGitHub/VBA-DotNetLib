Attribute VB_Name = "StringEqualsExample2"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 30, 2023
'@LastModified December 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.equals?view=netframework-4.8.1#system-string-equals(system-string)

Option Explicit

''
' The following example demonstrates the Equals method. It compares the
' title-cased word "File" with an equivalent word, its lowercase equivalent,
' its uppercase equivalent, and a word that contains LATIN SMALL LETTER DOTLESS
' I (U+0131) instead of LATIN SMALL LETTER I (U+0069). Because the Equals(String)
' method performs an ordinal comparison, only the comparison with an identical
' word returns true.
''
Public Sub StringEqualsExample2()
    Dim pvtWord As DotNetLib.String
    Set pvtWord = Strings.Create("File")
    Dim others() As DotNetLib.String
    Call ArrayEx.CreateInitialize1D(others, pvtWord.ToLower(), pvtWord, pvtWord.ToUpper(), Strings.CreateUnescape("F\u0131le"))
    Dim varOther As Variant
    For Each varOther In others
        Dim other As DotNetLib.String
        Set other = varOther
        If (pvtWord.Equals(other)) Then
            Debug.Print VBAString.Format("{0} = {1}", pvtWord, other)
        Else
            Debug.Print VBAString.Format("{0} {1} {2}", pvtWord, Regex.Unescape("\u2260"), other)
        End If
    Next
End Sub

' The example displays the following output:
'       File ? file
'       File = File
'       File ? FILE
'       File ? File
