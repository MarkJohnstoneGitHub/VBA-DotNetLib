Attribute VB_Name = "UriUnescapeDataStringExample"
'@Folder "Examples.System.Uri.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 26, 2023
'@LastModified January 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.unescapedatastring?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example unescapes a URI, and then converts any plus
' characters ("+") into spaces.
''
Public Sub UriUnescapeDataStringExample()
    Dim DataString As String
    DataString = Uri.UnescapeDataString(".NET+Framework")
    Debug.Print VBString.Format("Unescaped string: {0}", DataString)

    Dim PlusString As String
    PlusString = Replace(DataString, "+", " ")
    Debug.Print VBString.Format("plus to space string: {0}", PlusString)
End Sub

' Output
'    Unescaped string: .NET+Framework
'    plus to space string: .NET Framework

