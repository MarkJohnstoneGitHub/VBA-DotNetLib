Attribute VB_Name = "StringBuilderCapacityExample"
'@Folder("Examples.System.Text.StringBuilder.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 30, 2023
'@LastModified October 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.capacity?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates the Capacity property.
''
Public Sub StringBuilderCapacity()
    Dim sb1 As DotNetLib.StringBuilder
    Set sb1 = StringBuilder.Create("abc")
    Dim sb2 As DotNetLib.StringBuilder
    Set sb2 = StringBuilder.Create("abc", 16)
    
    Debug.Print
    Debug.Print VBAString.Format("a1) sb1.Length = {0}, sb1.Capacity = {1}", sb1.length, sb1.Capacity)
    Debug.Print VBAString.Format("a2) sb2.Length = {0}, sb2.Capacity = {1}", sb2.length, sb2.Capacity)
    Debug.Print VBAString.Format("a3) sb1.ToString() = ""{0}"", sb2.ToString() = ""{1}""", _
                                sb1.ToString(), sb2.ToString())
    Debug.Print VBAString.Format("a4) sb1 equals sb2: {0}", sb1.Equals(sb2))
    
    Debug.Print
    Debug.Print "Ensure sb1 has a capacity of at least 50 characters."
    Call sb1.EnsureCapacity(50)
    
    Debug.Print
    Debug.Print VBAString.Format("b1) sb1.Length = {0}, sb1.Capacity = {1}", sb1.length, sb1.Capacity)
    Debug.Print VBAString.Format("b2) sb2.Length = {0}, sb2.Capacity = {1}", sb2.length, sb2.Capacity)
    Debug.Print VBAString.Format("b3) sb1.ToString() = ""{0}"", sb2.ToString() = ""{1}""", _
                                sb1.ToString(), sb2.ToString())
    Debug.Print VBAString.Format("b4) sb1 equals sb2: {0}", sb1.Equals(sb2))
    
    Debug.Print
    Debug.Print "Set the length of sb1 to zero."
    Debug.Print "Set the capacity of sb2 to 51 characters."
    sb1.length = 0
    sb2.Capacity = 51
    
    Debug.Print
    Debug.Print VBAString.Format("c1) sb1.Length = {0}, sb1.Capacity = {1}", sb1.length, sb1.Capacity)
    Debug.Print VBAString.Format("c2) sb2.Length = {0}, sb2.Capacity = {1}", sb2.length, sb2.Capacity)
    Debug.Print VBAString.Format("c3) sb1.ToString() = ""{0}"", sb2.ToString() = ""{1}""", _
                                sb1.ToString(), sb2.ToString())
    Debug.Print VBAString.Format("c4) sb1 equals sb2: {0}", sb1.Equals(sb2))
End Sub

'/*
'The example displays the following output:
'
'a1) sb1.Length = 3, sb1.Capacity = 16
'a2) sb2.Length = 3, sb2.Capacity = 16
'a3) sb1.ToString() = "abc", sb2.ToString() = "abc"
'a4) sb1 equals sb2: True
'
'Ensure sb1 has a capacity of at least 50 characters.
'
'b1) sb1.Length = 3, sb1.Capacity = 50
'b2) sb2.Length = 3, sb2.Capacity = 16
'b3) sb1.ToString() = "abc", sb2.ToString() = "abc"
'b4) sb1 equals sb2: False
'
'Set the length of sb1 to zero.
'Set the capacity of sb2 to 51 characters.
'
'c1) sb1.Length = 0, sb1.Capacity = 50
'c2) sb2.Length = 3, sb2.Capacity = 51
'c3) sb1.ToString() = "", sb2.ToString() = "abc"
'c4) sb1 equals sb2: False
'*/


