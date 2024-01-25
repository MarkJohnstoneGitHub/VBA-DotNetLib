Attribute VB_Name = "UriCheckHostNameExample"
'@Folder("Examples.System.Uri.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 25, 2023
'@LastModified January 25, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.checkhostname?view=netframework-4.8.1#examples

Option Explicit

''
' The following example checks whether the host name is valid.
''
Public Sub UriCheckHostNameExample()
    Debug.Print UriHostNameTypeHelper.ToString(Uri.CheckHostName("www.contoso.com"))
End Sub

' Output
'   Dns
