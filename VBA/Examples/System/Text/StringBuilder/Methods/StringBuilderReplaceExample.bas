Attribute VB_Name = "StringBuilderReplaceExample"
'@Folder("Examples.System.Text.StringBuilder.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 31, 2023
'@LastModified October 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.replace?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates the Replace method.
''
Public Sub StringBuilderReplace()
    '      0----+----1----+----2----+----3----+----4---
    '      01234567890123456789012345678901234567890123
    Dim str As String
    str = "The quick br!wn d#g jumps #ver the lazy cat."
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create(str)

    Debug.Print
    Debug.Print "StringBuilder.Replace method"
    Debug.Print

    Debug.Print "Original value:"
    Call Show(sb)

    Call sb.Replace_2("#", "!", 15, 29)        ' Some '#' -> '!'
    Call Show(sb)
    Call sb.Replace("!", "o")                 ' All '!' -> 'o'
    Call Show(sb)
    Call sb.Replace("cat", "dog")             ' All "cat" -> "dog"
    Call Show(sb)
    Call sb.Replace_2("dog", "fox", 15, 20)    ' Some "dog" -> "fox"
    
    Debug.Print "Final value:"
    Call Show(sb)
End Sub

Private Sub Show(ByVal sbs As DotNetLib.StringBuilder)
    Dim rule1 As String
    rule1 = "0----+----1----+----2----+----3----+----4---"
    Dim rule2 As String
    rule2 = "01234567890123456789012345678901234567890123"

    Debug.Print rule1
    Debug.Print rule2
    Debug.Print VBAString.Format("{0}", sbs.ToString())
    Debug.Print
End Sub

'/*
'This example produces the following results:
'
'StringBuilder.Replace Method
'
'Original value:
'0----+----1----+----2----+----3----+----4---
'01234567890123456789012345678901234567890123
'The quick br!wn d#g jumps #ver the lazy cat.
'
'0----+----1----+----2----+----3----+----4---
'01234567890123456789012345678901234567890123
'The quick br!wn d!g jumps !ver the lazy cat.
'
'0----+----1----+----2----+----3----+----4---
'01234567890123456789012345678901234567890123
'The quick brown dog jumps over the lazy cat.
'
'0----+----1----+----2----+----3----+----4---
'01234567890123456789012345678901234567890123
'The quick brown dog jumps over the lazy dog.
'
'Final value:
'0----+----1----+----2----+----3----+----4---
'01234567890123456789012345678901234567890123
'The quick brown fox jumps over the lazy dog.
'
'*/
