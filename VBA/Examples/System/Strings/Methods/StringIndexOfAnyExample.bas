Attribute VB_Name = "StringIndexOfAnyExample"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 31, 2023
'@LastModified December 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.indexofany?view=netframework-4.8.1#system-string-indexofany(system-char())

Option Explicit

''
' The following example finds the first vowel in a string.
''
Public Sub StringIndexOfAnyExample()
    Dim chars As DotNetLib.String
    Set chars = Strings.Concat11("a", "e", "i", "o", "u", "y", _
                    "A", "E", "I", "O", "U", "Y")
    Dim s As DotNetLib.String
    Set s = Strings.Create("The long and winding road...")
    Debug.Print VBString.Format(Regex.Unescape("The first vowel in \n   {0}\nis found at position {1}"), _
                        s, s.IndexOfAny(chars) + 1)
End Sub

' The example displays the following output:
'       The first vowel in
'          The long and winding road...
'       is found at position 3


