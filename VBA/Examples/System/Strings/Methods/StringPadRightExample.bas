Attribute VB_Name = "StringPadRightExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 4, 2024
'@LastModified January 4, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.padright?view=netframework-4.8.1#system-string-padright(system-int32)

Option Explicit

''
' The following example demonstrates the PadRight method.
''
Public Sub StringPadRightExample()
    Dim str As DotNetLib.String
    Set str = Strings.Create("BBQ and Slaw")
    
    Debug.Print "|";
    Debug.Print str.PadRight(15).ToString();
    Debug.Print "|"       ' Displays "|BBQ and Slaw   |".

    Debug.Print "|";
    Debug.Print str.PadRight(5).ToString();
    Debug.Print "|"       ' Displays "|BBQ and Slaw|".
End Sub
