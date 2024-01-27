Attribute VB_Name = "StringBuilderAppendExample"
'@Folder "Examples.System.Text.StringBuilder.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 31, 2023
'@LastModified January 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.appendformat?view=netframework-4.8.1

Option Explicit

Public Sub StringBuilderAppend()
    Dim flag As Boolean
    flag = False
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create()
    
    Call sb.Append("The value of the flag is ").Append3(flag).Append(".")
    Debug.Print sb.ToString()
End Sub

' The example displays the following output:
'       The value of the flag is False.
