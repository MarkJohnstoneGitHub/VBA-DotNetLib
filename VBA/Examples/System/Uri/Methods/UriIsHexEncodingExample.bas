Attribute VB_Name = "UriIsHexEncodingExample"
'@Folder "Examples.System.Uri.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 26, 2023
'@LastModified January 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.ishexencoding?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example determines whether a character is hexadecimal
' encoded and, if so, writes the equivalent character to the console.
''
Public Sub UriIsHexEncodingExample()
    Dim testString As String
    testString = "%75"
    Dim pvtIndex As Long
    pvtIndex = 0
    If (Uri.IsHexEncoding(testString, pvtIndex)) Then
        Debug.Print VBString.Format("The character is {0}", Uri.HexUnescape(testString, pvtIndex))
    Else
        Debug.Print "The character is not hexadecimal encoded"
    End If
End Sub

' Output
'    The character Is u

