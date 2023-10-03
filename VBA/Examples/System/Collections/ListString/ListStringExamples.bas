Attribute VB_Name = "ListStringExamples"
'@Folder("Examples.System.Collections.ListString")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 3, 2023
'@LastModified October 4, 2023

'Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1.contains?view=netframework-4.8.1

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
    
    names.Insert 4, "Elizabeth"
    Debug.Print "List after : names.Insert 4, 'Elizabeth'"
    pvtIndex = 0
    For Each varName In names
        Debug.Print Strings.Format("[{0}] {1}", pvtIndex, varName)
        pvtIndex = pvtIndex + 1
    Next
    Debug.Print
    
    Dim pvtName As String
    
    Debug.Print "names.Contains(item)"
    Debug.Print Strings.Format("Contains 'Bob' : {0}", names.Contains("Bob"))
    Debug.Print Strings.Format("Contains 'Luke' : {0}", names.Contains("Luke"))
    Debug.Print Strings.Format("Contains 'Susan' : {0}", names.Contains("Susan"))
    Debug.Print Strings.Format("Contains 'Elizabeth' : {0}", names.Contains("Elizabeth"))
    Debug.Print
    
    
    Debug.Print "names.IndexOf(item)"
    pvtName = "James"
    Debug.Print Strings.Format("{0,3} : List.IndexOf('{1}')", names.IndexOf(pvtName), pvtName)
    pvtName = "Mark"
    Debug.Print Strings.Format("{0,3} : List.IndexOf('{1}')", names.IndexOf(pvtName), pvtName)
    pvtName = "Brian"
    Debug.Print Strings.Format("{0,3} : List.IndexOf('{1}')", names.IndexOf(pvtName), pvtName)
    pvtName = "Luke"
    Debug.Print Strings.Format("{0,3} : List.IndexOf('{1}')", names.IndexOf(pvtName), pvtName)
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
    Debug.Print "names.BinarySearch(item)"
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
'    names.BinarySearch (Item)
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

