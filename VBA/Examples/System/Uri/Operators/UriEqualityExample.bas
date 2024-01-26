Attribute VB_Name = "UriEqualityExample"
'@Folder("Examples.System.Uri.Operators")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 26, 2023
'@LastModified January 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.op_equality?view=netframework-4.8.1#examples

Option Explicit

''
' This example creates three Uri instances from strings and compares them to
' determine whether they represent the same value. Address1 and Address2 are
' the same because the Fragment portion is ignored for this comparison. The
' outcome is written to the console.
''
Public Sub UriEqualityExample()
    ' Create some Uris.
    Dim address1 As DotNetLib.Uri
    Set address1 = Uri.Create("http://www.contoso.com/index.htm#search")
    Dim address2 As DotNetLib.Uri
    Set address2 = Uri.Create("http://www.contoso.com/index.htm")
    Dim address3 As DotNetLib.Uri
    Set address3 = Uri.Create("http://www.contoso.com/index.htm?date=today")
    
    ' The first two are equal because the fragment is ignored.
    If Uri.Equality(address1, address2) Then
        Debug.Print VBString.Format("{0} is equal to {1}", address1.ToString(), address2.ToString())
    End If
    
    ' The second two are not equal.
    If Uri.Inequality(address2, address3) Then
        Debug.Print VBString.Format("{0} is not equal to {1}", address2.ToString(), address3.ToString())
    End If
End Sub

' Output
' http://www.contoso.com/index.htm#search is equal to http://www.contoso.com/index.htm
' http://www.contoso.com/index.htm is not equal to http://www.contoso.com/index.htm?date=today
