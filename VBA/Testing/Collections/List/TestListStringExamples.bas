Attribute VB_Name = "TestListStringExamples"
'@Folder("Testing.Collections.List")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 3, 2023
'@LastModified October 3, 2023

Option Explicit

'Test example using ListString i.e. List<String>
Private Sub TestListStringExample1()
    Dim names As DotNetLib.ListString
    Set names = ListString.Create()

    names.Add "Bob"
    names.Add "Mary"
    names.Add "James"
    names.Add "Mark"
    names.Add "Brian"
    names.Add "susan"
    
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
