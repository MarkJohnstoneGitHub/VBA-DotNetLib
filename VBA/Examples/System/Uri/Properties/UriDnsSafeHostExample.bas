Attribute VB_Name = "UriDnsSafeHostExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 22, 2023
'@LastModified January 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.dnssafehost?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates a Uri instance from a string. It illustrates
' the difference between the value returned from Host, which returns the host
' name or address specified in the URI, and the value returned from DnsSafeHost,
' which returns an address that is safe to use in DNS resolution.
''
Public Sub UriDnsSafeHostExample()
    ' Create new Uri using a string address.
    Dim address As DotNetLib.Uri
    Set address = Uri.Create("http://[fe80::200:39ff:fe36:1a2d%254]/temp/example.htm")

    ' Make the address DNS safe.

    ' The following outputs "[fe80::200:39ff:fe36:1a2d]".
    Debug.Print address.Host

    ' The following outputs "fe80::200:39ff:fe36:1a2d%254".
    Debug.Print address.DnsSafeHost
End Sub
