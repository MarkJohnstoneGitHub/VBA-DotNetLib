Attribute VB_Name = "StringCharCodesExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 3, 2024
'@LastModified January 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.chars?view=netframework-4.8.1#remarks

'@Notes Currently not implementing Char object instead returning a Char converted to AscW i.e Long.

Option Explicit

''
' The index parameter is zero-based.
''
Public Sub StringCharCodesExample()
    Dim str1 As DotNetLib.String
    Set str1 = Strings.Create("Test")
    Dim ctr As Long
    For ctr = 0 To str1.length - 1
        Debug.Print VBAString.Format("{0} ", ChrW$(str1.CharCodes(ctr)));
    Next
    Debug.Print
    For ctr = 0 To str1.length - 1
        Debug.Print VBAString.Format("{0} ", str1.CharCodes(ctr));
    Next
End Sub

' The example displays the following output:
' T e s t
' 84 101 115 116
