Attribute VB_Name = "UriEqualsExample"
'@Folder "Examples.System.Uri.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 25, 2023
'@LastModified January 25, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.equals?view=netframework-4.8.1#examples

Option Explicit

''
' This example creates two Uri instances from strings and compares them to
' determine whether they represent the same value. address1 and address2
' are the same because the Fragment portion is ignored for this comparison.
' The outcome is written to the console.
''
Public Sub UriEqualsExample()
    ' Create some Uris.
    Dim address1 As DotNetLib.Uri
    Set address1 = Uri.Create("http://www.contoso.com/index.htm#search")
    Dim address2 As DotNetLib.Uri
    Set address2 = Uri.Create("http://www.contoso.com/index.htm")
    If (address1.Equals(address2)) Then
        Debug.Print "The two addresses are equal"
    Else
        Debug.Print "The two addresses are not equal"
    End If
    ' Will output "The two addresses are equal"
End Sub

' Output
'   The two addresses are equal

