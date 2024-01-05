Attribute VB_Name = "StringSplitExample5"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 5, 2024
'@LastModified January 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.split?view=netframework-4.8.1#system-string-split(system-char()-system-stringsplitoptions)

Option Explicit

''
' If the separator parameter is null or contains no characters, white-space
' characters are assumed to be the delimiters. White-space characters are
' defined by the Unicode standard and the Char.IsWhiteSpace method returns
' true if they are passed to it.
''
Public Sub StringSplitExample5()
    Dim phrase As DotNetLib.String
    Set phrase = Strings.Create("The quick  brown fox")
    
    Dim subs() As String
    subs = phrase.Split(VBA.vbNullString, StringSplitOptions.StringSplitOptions_RemoveEmptyEntries)
    Dim varSub As Variant
    For Each varSub In subs
        Debug.Print VBAString.Format("Substring: {0}", varSub)
    Next
    Debug.Print
    subs = phrase.Split(Empty, StringSplitOptions.StringSplitOptions_RemoveEmptyEntries)
    For Each varSub In subs
        Debug.Print VBAString.Format("Substring: {0}", varSub)
    Next
End Sub

'Output:
'Substring: The
'Substring: quick
'Substring: brown
'Substring: fox
'
'Substring: The
'Substring: quick
'Substring: brown
'Substring: fox
