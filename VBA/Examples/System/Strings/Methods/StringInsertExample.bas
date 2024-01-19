Attribute VB_Name = "StringInsertExample"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 1, 2024
'@LastModified January 1, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.insert?view=netframework-4.8.1#examples

Option Explicit

''
' The following example inserts a space character in the fourth character
' position (the character at index 3) of a string.
''
Public Sub StringInsertExample()
    Dim original As DotNetLib.String
    Set original = Strings.Create("aaabbb")
    Debug.Print VBString.Format("The original string: '{0}'", original)
    Dim modified As DotNetLib.String
    Set modified = original.Insert2(3, " ")
    Debug.Print VBString.Format("The modified string: '{0}'", modified)
End Sub

' The example displays the following output:
'     The original string: 'aaabbb'
'     The modified string: 'aaa bbb'
