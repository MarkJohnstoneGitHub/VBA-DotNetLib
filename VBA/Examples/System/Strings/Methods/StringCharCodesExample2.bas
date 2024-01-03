Attribute VB_Name = "StringCharCodesExample2"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 3, 2024
'@LastModified January 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.chars?view=netframework-4.8.1#remarks

Option Explicit

''
' The index parameter is zero-based.
' This property returns the AscW long of a Char Object at the position specified
' by the index parameter.  However, a Unicode character might be represented by
' more than one Char. Use the System.Globalization.StringInfo class to work with
' Unicode characters instead of Char objects.
'
' This example iterates a string by characters checking if the character is a surrogate.
''
Public Sub StringCharCodesExample2()
    Dim ctr As Long
    Dim s1 As DotNetLib.String
    Set s1 = Strings.CreateUnescape("a\uD800\uDC00y")
    
    Debug.Print VBAString.Format("String '{0}' {1}", s1, IIf(s1.IsSurrogate(), "contains surrogate pairs", ""))
    For ctr = 0 To s1.length - 1
        Dim pvtCharCode As Long
        pvtCharCode = s1.CharCodes(ctr)
        Debug.Print VBAString.Format("{0} {1} {2}", ChrW$(pvtCharCode), pvtCharCode, IIf(Char.IsSurrogate2(pvtCharCode), "Is Surrogate", ""))
    Next
End Sub

'The example displays the following output:
'String 'a??y' contains surrogate pairs
'a 97
'Print 55296 Is Surrogate
'Print 56320 Is Surrogate
'y 121

