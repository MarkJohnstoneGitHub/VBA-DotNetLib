Attribute VB_Name = "StringBuilderCharsExample"
'@Folder "Examples.System.Text.StringBuilder.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 30, 2023
'@LastModified October 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.chars?view=netframework-4.8.1#remarks

Option Explicit

''
' Counts the number of alphabetic, white-space, and punctuation characters
' in a string.
''
Public Sub StringBuilderChars()
    Dim nAlphabeticChars As Long
    nAlphabeticChars = 0
    Dim nWhitespace As Long
    nWhitespace = 0
    Dim nPunctuation As Long
    nPunctuation = 0
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create("This is a simple sentence.")
    
    Dim ctr As Long
    For ctr = 0 To sb.Length - 1
        Dim ch As String
        ch = sb.item(ctr)
        If (Char.IsLetter(ch, 0)) Then
            nAlphabeticChars = nAlphabeticChars + 1
        ElseIf (Char.IsWhiteSpace(ch, 0)) Then
            nWhitespace = nWhitespace + 1
        ElseIf (Char.IsPunctuation(ch, 0)) Then
            nPunctuation = nPunctuation + 1
        End If
    Next
    
    Debug.Print VBString.Format("The sentence '{0}' has:", sb)
    Debug.Print VBString.Format("   Alphabetic characters: {0}", nAlphabeticChars)
    Debug.Print VBString.Format("   White-space characters: {0}", nWhitespace)
    Debug.Print VBString.Format("   Punctuation characters: {0}", nPunctuation)
End Sub

' The example displays the following output:
'       The sentence 'This is a simple sentence.' has:
'          Alphabetic characters: 21
'          White-space characters: 4
'          Punctuation characters: 1
