Attribute VB_Name = "UriFromHexExample"
'@Folder "Examples.System.Uri.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 25, 2023
'@LastModified January 25, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.fromhex?view=netframework-4.8.1#examples

Option Explicit

''
' The following example determines whether a character is a hexadecimal character
' and, if it is, writes the corresponding decimal value to the console.
''
Public Sub UriFromHexExample()
    Dim testchar As String
    testchar = "e"
    If (Uri.IsHexDigit(testchar) = True) Then
        Debug.Print VBString.Format("'{0}' is the hexadecimal representation of {1}", testchar, Uri.FromHex(testchar))
    Else
        Debug.Print VBString.Format("'{0}' is not a hexadecimal character", testchar)
    End If

    Dim returnString As String
    returnString = Uri.HexEscape(testchar)
    Debug.Print VBString.Format("The hexadecimal value of '{0}' is {1}", testchar, returnString)
End Sub

' Output
'    'e' is the hexadecimal representation of 14
'    The hexadecimal value of 'e' is %65
