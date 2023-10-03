Attribute VB_Name = "ListStringExamples"
'@Folder("Examples.System.Collections.ListString")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 3, 2023
'@LastModified October 3, 2023

Option Explicit

'Test example using ListString i.e. List<String>
Private Sub ListStringExample1()
    Dim names As DotNetLib.ListString
    Set names = ListString.Create()
    names.Add "Bob"
    names.Add "Mary"
    names.Add "James"
    names.Add "Mark"
    names.Add "Brian"
    names.Add "susan"
    names.Add "Michael"
    
    If names.Contains("Bob") Then
        Debug.Print "List contains 'Bob'"
    End If
    Debug.Print
    
    Debug.Print "Names list as added:"
    Dim varName As Variant
    For Each varName In names
        Debug.Print varName
    Next
    Debug.Print

    Debug.Print "Names list sorted:"
    names.Sort
    For Each varName In names
        Debug.Print varName
    Next
    Debug.Print
    
    Debug.Print "Names list reversed sorted list:"
    names.Reverse
    For Each varName In names
        Debug.Print varName
    Next
End Sub

' Output:
'
'    List Contains 'Bob'
'
'    Names list as added:
'    Bob
'    Mary
'    James
'    Mark
'    Brian
'    susan
'    Michael
'
'    Names list sorted:
'    Bob
'    Brian
'    James
'    Mark
'    Mary
'    Michael
'    susan
'
'    Names list reversed sorted list:
'    susan
'    Michael
'    Mary
'    Mark
'    James
'    Brian
'    Bob




