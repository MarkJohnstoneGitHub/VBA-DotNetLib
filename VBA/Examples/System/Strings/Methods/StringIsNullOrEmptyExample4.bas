Attribute VB_Name = "StringIsNullOrEmptyExample4"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 2, 2024
'@LastModified January 2, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.isnullorempty?view=netframework-4.8.1#what-is-an-empty-string

Option Explicit

''
' A string is empty if it is explicitly assigned an empty string ("") or
' String.Empty. An empty string has a Length of 0. The following example
' creates an empty string and displays its value and its length.
''
Public Sub StringIsNullOrEmptyExample4()
    Dim s As DotNetLib.String
    Set s = Strings.Create("")
    Debug.Print VBAString.Format("The length of '{0}' is {1}.", s, s.length)
End Sub

' The example displays the following output:
'       The length of '' is 0.
