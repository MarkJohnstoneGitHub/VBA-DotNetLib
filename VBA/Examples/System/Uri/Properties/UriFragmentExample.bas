Attribute VB_Name = "UriFragmentExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 23, 2023
'@LastModified January 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.fragment?view=netframework-4.8.1#examples

Option Explicit

Public Sub UriFragmentExample()
    '// Create Uri
    Dim uriAddress As DotNetLib.Uri
    Set uriAddress = Uri.Create("http://www.contoso.com/index.htm#search")
    Debug.Print (uriAddress.Fragment)
    Debug.Print VBString.Format("Uri {0} the default port ", IIf(uriAddress.IsDefaultPort, "uses", "does not use"))
    
    Debug.Print VBString.Format("The path of this Uri is {0}", uriAddress.GetLeftPart(UriPartial.UriPartial_Path))
    Debug.Print VBString.Format("Hash code {0}", uriAddress.GetHashCode())
End Sub

' The example displays output similar to the following:
'        #search
'        Uri uses the default port
'        The path of this Uri is http://www.contoso.com/index.htm
'        Hash code -988419291
