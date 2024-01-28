Attribute VB_Name = "StringBuilderAppend2Example"
'@Folder("Examples.System.Text.StringBuilder.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 28, 2023
'@LastModified January 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.append?view=netframework-4.8.1#system-text-stringbuilder-append(system-string-system-int32-system-int32)

Option Explicit

''
' The Append(String, Int32, Int32) method modifies the existing instance of
' this class; it does not return a new class instance. Because of this, you
' can call a method or property on the existing reference and you do not have
' to assign the return value to a StringBuilder object, as the following
' example illustrates.
''
Public Sub StringBuilderAppend2Example()
    Dim str As DotNetLib.String
    Set str = Strings.Create("First;George Washington;1789;1797")
    Dim pvtIndex As Long
    pvtIndex = 0
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create()
    Dim pvtLength As Long
    pvtLength = str.IndexOf9(";", pvtIndex)
    Call sb.Append2(str.ToString, pvtIndex, pvtLength).Append(" President of the United States: ")
    pvtIndex = pvtIndex + pvtLength + 1
    pvtLength = str.IndexOf9(";", pvtIndex) - pvtIndex
    Call sb.Append2(str.ToString, pvtIndex, pvtLength).Append(", from ")
    pvtIndex = pvtIndex + pvtLength + 1
    pvtLength = str.IndexOf9(";", pvtIndex) - pvtIndex
    Call sb.Append2(str.ToString, pvtIndex, pvtLength).Append(" to ")
    pvtIndex = pvtIndex + pvtLength + 1
    Call sb.Append2(str.ToString, pvtIndex, str.Length - pvtIndex)
    Debug.Print sb.ToString()
End Sub

' The example displays the following output:
'    First President of the United States: George Washington, from 1789 to 1797
