Attribute VB_Name = "StringSplitExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 5, 2024
'@LastModified January 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.split?view=netframework-4.8.1#example

Option Explicit

''
' Calls the Split(Char[]) overload and passes in a single delimiter.
''
Public Sub StringSplitExample()
    Dim s As DotNetLib.String
    Set s = Strings.Create("You win some. You lose some.")

    Dim subs() As String
    subs = s.Split(" ")
    
    Dim varSub As Variant
    For Each varSub In subs
        Debug.Print VBAString.Format("Substring: {0}", varSub)
    Next
End Sub

' This example produces the following output:
'
' Substring: You
' Substring: win
' Substring: some.
' Substring: You
' Substring: lose
' Substring: some.
