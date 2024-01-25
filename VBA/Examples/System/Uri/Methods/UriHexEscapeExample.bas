Attribute VB_Name = "UriHexEscapeExample"
'@Folder("Examples.System.Uri.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 26, 2023
'@LastModified January 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.hexescape?view=netframework-4.8.1#examples

Option Explicit

''
' The following example converts a character to its hexadecimal equivalent and
' writes it to the console.
''
Public Sub UriHexEscapeExample()
    Dim testChar As String
    testChar = "e"
    If (Uri.IsHexDigit(testChar) = True) Then
        Debug.Print VBString.Format("'{0}' is the hexadecimal representation of {1}", testChar, Uri.FromHex(testChar))
    Else
        Debug.Print VBString.Format("'{0}' is not a hexadecimal character", testChar)
    End If
    Dim returnString As String
    returnString = Uri.HexEscape(testChar)
    Debug.Print VBString.Format("The hexadecimal value of '{0}' is {1}", testChar, returnString)
End Sub

' Output
'    'e' is the hexadecimal representation of 14
'    The hexadecimal value of 'e' is %65
