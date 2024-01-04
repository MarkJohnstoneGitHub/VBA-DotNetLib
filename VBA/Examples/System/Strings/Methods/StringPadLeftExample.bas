Attribute VB_Name = "StringPadLeftExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 4, 2024
'@LastModified January 4, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.padleft?view=netframework-4.8.1#system-string-padleft(system-int32)

Option Explicit

''
' The following example demonstrates the PadLeft method.
''
Public Sub StringPadLeftExample()
    Dim str As DotNetLib.String
    Set str = Strings.Create("BBQ and Slaw")
    Debug.Print str.PadLeft(15).ToString()  ' Displays "   BBQ and Slaw".
    Debug.Print str.PadLeft(5).ToString()   ' Displays "BBQ and Slaw".
End Sub
