Attribute VB_Name = "StringSplitExample2"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 5, 2024
'@LastModified January 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.split?view=netframework-4.8.1#system-string-split(system-char())

Option Explicit

''
' The following example demonstrates how to extract individual words from a block
' of text by treating the space character ( ) and tab character (\t) as delimiters.
' The string being split includes both of these characters.
''
Public Sub StringSplitExample2()
    Dim s As DotNetLib.String
    Set s = Strings.CreateUnescape("Today\tI'm going to school")
    
    Dim subs() As String
    subs = s.Split(" " & VBAString.Unescape("\t"))
    
    Dim varSub As Variant
    For Each varSub In subs
        Debug.Print VBAString.Format("Substring: {0}", varSub)
    Next
End Sub

' This example produces the following output:
'
' Substring: Today
' Substring: I'm
' Substring: going
' Substring: to
' Substring: school
