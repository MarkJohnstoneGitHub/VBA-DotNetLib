Attribute VB_Name = "StringBuilderClearExample"
'@Folder("Examples.System.Text.StringBuilder.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 31, 2023
'@LastModified October 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.clear?view=netframework-4.8.1#examples

Option Explicit

''
' The following example instantiates a StringBuilder object with a string,
' calls the Clear method, and then appends a new string.
''
Public Sub StringBuilderClear()
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create("This is a string.")
    
    Debug.Print VBAString.Format("{0} ({1} characters)", sb.ToString(), sb.length)

    Call sb.Clear
    Debug.Print VBAString.Format("{0} ({1} characters)", sb.ToString(), sb.length)

    Call sb.Append("This is a second string.")
    Debug.Print VBAString.Format("{0} ({1} characters)", sb.ToString(), sb.length)
End Sub

' The example displays the following output:
'       This is a string. (17 characters)
'        (0 characters)
'       This is a second string. (24 characters)
