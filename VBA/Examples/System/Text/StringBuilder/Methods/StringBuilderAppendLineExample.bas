Attribute VB_Name = "StringBuilderAppendLineExample"
'@Folder "Examples.System.Text.StringBuilder.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 31, 2023
'@LastModified January 27, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.appendline?view=netframework-4.8.1

Option Explicit

''
' The following example demonstrates the AppendLine method.
''
Public Sub StringBuilderAppendLine()
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create()
    Dim line As String
    line = "A line of text."
    Dim pvtNumber As Long
    pvtNumber = 123

    ' Append two lines of text.
    Call sb.AppendLine2("The first line of text.")
    Call sb.AppendLine2(line)
    
    ' Append a new line, an empty string, and a null cast as a string.
    Call sb.AppendLine
    Call sb.AppendLine2(VBA.vbNullString)
    Call sb.AppendLine2(Empty)
    
    ' Append the non-string value, 123, and two new lines.
    Call sb.Append6(pvtNumber).AppendLine().AppendLine

    ' Append two lines of text.
    Call sb.AppendLine2(line)
    Call sb.AppendLine2("The last line of text.")

    ' Convert the value of the StringBuilder to a string and display the string.
    Debug.Print sb.ToString()
End Sub

'/*
'This example produces the following results:
'
'The first line of text.
'A line of text.
'
'
'
'123
'
'A line of text.
'The last line of text.
'*/
