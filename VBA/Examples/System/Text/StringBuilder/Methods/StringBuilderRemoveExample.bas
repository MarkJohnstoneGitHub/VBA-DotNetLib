Attribute VB_Name = "StringBuilderRemoveExample"
'@Folder "Examples.System.Text.StringBuilder.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 31, 2023
'@LastModified October 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.remove?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates the Remove method.
''
Public Sub StringBuilderRemove()
    Dim rule1 As String
    rule1 = "0----+----1----+----2----+----3----+----4---"
    Dim rule2 As String
    rule2 = "01234567890123456789012345678901234567890123"
    Dim str As String
    str = "The quick brown fox jumps over the lazy dog."
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create(str)
    
    Debug.Print
    Debug.Print "StringBuilder.Remove method"
    Debug.Print
    Debug.Print "Original value:"
    Debug.Print rule1
    Debug.Print rule2
    Debug.Print VBString.Format("{0}", sb.ToString())
    Debug.Print
    
    Call sb.Remove(10, 6) ' Remove "brown "

    Debug.Print "New value:"
    Debug.Print rule1
    Debug.Print rule2
    Debug.Print VBString.Format("{0}", sb.ToString())
End Sub

'/*
'This example produces the following results:
'
'StringBuilder.Remove Method
'
'Original value:
'0----+----1----+----2----+----3----+----4---
'01234567890123456789012345678901234567890123
'The quick brown fox jumps over the lazy dog.
'
'New value:
'0----+----1----+----2----+----3----+----4---
'01234567890123456789012345678901234567890123
'The quick fox jumps over the lazy dog.
'
'*/
