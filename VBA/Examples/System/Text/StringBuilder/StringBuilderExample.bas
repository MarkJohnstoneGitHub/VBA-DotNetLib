Attribute VB_Name = "StringBuilderExample"
'@Folder("Examples.System.Text.StringBuilder")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 30, 2023
'@LastModified October 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=netframework-4.8.1

Option Explicit

''
' The following example shows how to call many of the methods defined by the
' StringBuilder class.
''
Public Sub StringBuilderExample()
    ' Create a StringBuilder that expects to hold 50 characters.
    ' Initialize the StringBuilder with "ABC".
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create("ABC", 50)

    ' Append three characters (D, E, and F) to the end of the StringBuilder.
    Call sb.Append("DEF")
    
    ' Append a format string to the end of the StringBuilder.
    Call sb.AppendFormat_3("GHI{0}{1}", "J", "k")
    
    ' Display the number of characters in the StringBuilder and its string.
    Debug.Print VBAString.Format("{0} chars: {1}", sb.length, sb.ToString())
    
    ' Insert a string at the beginning of the StringBuilder.
    Call sb.Insert(0, "Alphabet: ")
    
    ' Replace all lowercase k's with uppercase K's.
    Call sb.Replace("k", "K")
    
    ' Display the number of characters in the StringBuilder and its string.
    Debug.Print VBAString.Format("{0} chars: {1}", sb.length, sb.ToString())
End Sub

' This code produces the following output.
'
' 11 chars: ABCDEFGHIJk
' 21 chars: Alphabet: ABCDEFGHIJK


Public Sub StringBuilderExample2()
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create()
    Call sb.Append("This is the beginning of a sentence, ")
    Call sb.Replace("the beginning of ", "")
    
    'Call sb.Insert(sb.ToString(). .IndexOf("a ") + 2,  "complete ")
    Call sb.Replace(",", ".")
End Sub


'using System;
'using System.Text;
'
'public class Example
'{
'   public static void Main()
'   {
'      StringBuilder sb = new StringBuilder();
'      sb.Append("This is the beginning of a sentence, ");
'      sb.Replace("the beginning of ", "");
'      sb.Insert(sb.ToString().IndexOf("a ") + 2, "complete ");
'      sb.Replace(",", ".");
'      Console.WriteLine(sb.ToString());
'   }
'}
'// The example displays the following output:
'//        This is a complete sentence.
'
