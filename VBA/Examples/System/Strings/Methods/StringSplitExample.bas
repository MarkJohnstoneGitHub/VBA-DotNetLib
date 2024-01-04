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
' The following examples show three different overloads of String.Split().
' The first example calls the Split(Char[]) overload and passes in a single delimiter.
''
Public Sub StringSplitExampleA()
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

''
' As you can see, the period characters (.) are included in two of the substrings.
' If you want to exclude the period characters, you can add the period character
' as an additional delimiting character. The next example shows how to do this.
''
Public Sub StringSplitExampleB()
    Dim s As DotNetLib.String
    Set s = Strings.Create("You win some. You lose some.")
    Dim subs() As String
    subs = s.Split(" .")
    Dim varSub As Variant
    For Each varSub In subs
        Debug.Print VBAString.Format("Substring: {0}", varSub)
    Next
End Sub

' This example produces the following output:
'
' Substring: You
' Substring: win
' Substring: some
' Substring:
' Substring: You
' Substring: lose
' Substring: some
' Substring:

''
' The periods are gone from the substrings, but now two extra empty substrings
' have been included. These empty substring represent the substring between a
' word and the period that follows it. To omit empty substrings from the
' resulting array, you can call the Split(Char[], StringSplitOptions) overload
' and specify StringSplitOptions.RemoveEmptyEntries for the options parameter.
''
Public Sub StringSplitExampleC()
    Dim s As DotNetLib.String
    Set s = Strings.Create("You win some. You lose some.")
    Dim separators As String
    separators = " ."
    Dim subs() As String
    subs = s.Split(separators, StringSplitOptions_RemoveEmptyEntries)
    Dim varSub As Variant
    For Each varSub In subs
        Debug.Print VBAString.Format("Substring: {0}", varSub)
    Next
End Sub

' This example produces the following output:
'
' Substring: You
' Substring: win
' Substring: some
' Substring: You
' Substring: lose
' Substring: some
