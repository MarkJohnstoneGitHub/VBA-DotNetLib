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
    
    Debug.Print "Initial list:"
    Dim varName As Variant
    Dim pvtIndex As Long
    pvtIndex = 0
    For Each varName In names
        Debug.Print Strings.Format("[{0}] {1}", pvtIndex, varName)
        pvtIndex = pvtIndex + 1
    Next
    Debug.Print
    
    Debug.Print "names.Contains(item)"
    If names.Contains("Bob") Then
        Debug.Print "names.Contains : Contains 'Bob'"
    Else
        Debug.Print "names.Contains : Does not contain 'Bob'"
    End If
    
    If names.Contains("Luke") Then
        Debug.Print "names.Contains : Contains 'Luke'"
    Else
        Debug.Print "names.Contains : Does not contain 'Luke'"
    End If
    
    If names.Contains("Susan") Then
        Debug.Print "names.Contains : Contains 'Susan'"
    Else
        Debug.Print "names.Contains : Does not contain 'Susan'"
    End If
    Debug.Print
    
    Debug.Print "names.IndexOf(item)"
    Debug.Print names.IndexOf("James"); " : List.IndexOf('James')"
    Debug.Print names.IndexOf("Mark"); " : List.IndexOf('Mark')"
    Debug.Print names.IndexOf("Luke"); " : List.IndexOf('Luke')"
    Debug.Print
    
    Debug.Print "Sorted list:"
    names.Sort
    pvtIndex = 0
    For Each varName In names
        Debug.Print Strings.Format("[{0}] {1}", pvtIndex, varName)
        pvtIndex = pvtIndex + 1
    Next
    Debug.Print
    
    'Binary search on a sorted list
    'Searches the entire sorted List<T> for an element using the default comparer and returns the zero-based index of the element.
    'The List<T> must already be sorted according to the comparer implementation; otherwise, the result is incorrect.
    Debug.Print "list.BinarySearch(item)"
    Debug.Print names.BinarySearch("James"); " : List.BinarchSearch('James')"
    Debug.Print names.BinarySearch("susan"); " : List.BinarchSearch('susan')"
    Debug.Print names.BinarySearch("Greg"); " : List.BinarchSearch('Greg')"
    Debug.Print
    
    Debug.Print "Reverse list:"
    names.Reverse
    pvtIndex = 0
    For Each varName In names
        Debug.Print Strings.Format("[{0}] {1}", pvtIndex, varName)
        pvtIndex = pvtIndex + 1
    Next
End Sub

' Output:
'
'    Initial List:
'    [0] Bob
'    [1] Mary
'    [2] James
'    [3] Mark
'    [4] Brian
'    [5] susan
'    [6] Michael
'
'    names.Contains (Item)
'    names.Contains: Contains  'Bob'
'    names.Contains: Does Not contain  'Luke'
'    names.Contains: Does Not contain  'Susan'
'
'    names.IndexOf (Item)
'     2  : List.IndexOf('James')
'     3  : List.IndexOf('Mark')
'    -1  : List.IndexOf('Luke')
'
'    Sorted List:
'    [0] Bob
'    [1] Brian
'    [2] James
'    [3] Mark
'    [4] Mary
'    [5] Michael
'    [6] susan
'
'    List.BinarySearch (Item)
'     2  : List.BinarchSearch('James')
'     6  : List.BinarchSearch('susan')
'    -3  : List.BinarchSearch('Greg')
'
'    Reverse List:
'    [0] susan
'    [1] Michael
'    [2] Mary
'    [3] Mark
'    [4] James
'    [5] Brian
'    [6] Bob


