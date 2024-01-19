Attribute VB_Name = "StringPadRightExample2"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 4, 2024
'@LastModified January 4, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.padright?view=netframework-4.8.1#system-string-padright(system-int32-system-char)

Option Explicit

''
' The following example demonstrates the PadRight method.
''
Public Sub StringPadRightExample2()
    Dim str As DotNetLib.String
    Set str = Strings.Create("forty-two")
    Dim pad As String
    pad = "."
    Debug.Print str.PadRight(15, pad).ToString() ' Displays "forty-two......".
    Debug.Print str.PadRight(2, pad).ToString()  ' Displays "forty-two".
End Sub

