Attribute VB_Name = "StringLengthExample"
'@Folder("Examples.System.Strings.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 29, 2023
'@LastModified December 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.length?view=netframework-4.8.1#examples
Option Explicit

''
' The following example demonstrates the String.Length property.
''
Public Sub StringLengthExample()
    Dim str As DotNetLib.String
    Set str = Strings.Create("abcdefg")
    Debug.Print VBAString.Format("1) The length of '{0}' is {1}", str, str.length)
    Debug.Print VBAString.Format("2) The length of '{0}' is {1}", "xyz", Strings.Create("xyz").length)
    Dim pvtLength As Long
    pvtLength = str.length
    Debug.Print VBAString.Format("3) The length of '{0}' is {1}", str, pvtLength)
End Sub

' This example displays the following output:
'    1) The length of 'abcdefg' is 7
'    2) The length of 'xyz' is 3
'    3) The length of 'abcdefg' is 7

