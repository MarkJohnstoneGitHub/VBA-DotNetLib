Attribute VB_Name = "UriGetHashCodeExample"
'@Folder "Examples.System.Uri.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 26, 2023
'@LastModified January 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.gethashcode?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates a Uri instance and writes the hash code to the console.
''
Public Sub UriGetHashCodeExample()
    ' Create Uri
    Dim uriAddress As DotNetLib.Uri
    Set uriAddress = Uri.Create("http://www.contoso.com/index.htm#search")
    Debug.Print uriAddress.Fragment
    Debug.Print VBString.Format("Uri {0} the default port ", IIf(uriAddress.IsDefaultPort, "uses", "does not use"))
    
    Debug.Print VBString.Format("The path of this Uri is {0}", uriAddress.GetLeftPart(UriPartial.UriPartial_Path))
    Debug.Print VBString.Format("Hash code {0}", uriAddress.GetHashCode())
End Sub

' The example displays output similar to the following:
'        #search
'        Uri uses the default port
'        The path of this Uri is http://www.contoso.com/index.htm
'        Hash code -988419291
