Attribute VB_Name = "UriMakeRelativeUriExample"
'@Folder "Examples.System.Uri.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 26, 2023
'@LastModified January 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.makerelativeuri?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates 2 Uri instances. The difference in the path
' information is written to the console.
''
Public Sub UriMakeRelativeUriExample()
    ' Create a base Uri.
    Dim address1 As DotNetLib.Uri
    Set address1 = Uri.Create("http://www.contoso.com/")
    
    ' Create a new Uri from a string.
    Dim address2 As DotNetLib.Uri
    Set address2 = Uri.Create("http://www.contoso.com/index.htm?date=today")
    
    ' Determine the relative Uri.
    Debug.Print VBString.Format("The difference is {0}", address1.MakeRelativeUri(address2))
End Sub

' Output
'   The difference is index.htm?date=today

