Attribute VB_Name = "StringPadLeftExample2"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 4, 2024
'@LastModified January 4, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.padleft?view=netframework-4.8.1#system-string-padleft(system-int32-system-char)

Option Explicit

''
' The following example demonstrates the PadLeft method.
''
Public Sub StringPadLeftExample2()
    Dim str As DotNetLib.String
    Set str = Strings.Create("forty-two")
    Dim pad As String
    pad = "."
    Debug.Print str.PadLeft(15, pad).ToString()
    Debug.Print str.PadLeft(2, pad).ToString()
End Sub

' The example displays the following output:
'       ......forty-two
'       forty-two
