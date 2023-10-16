Attribute VB_Name = "ListStringExamples"
'@Folder("Examples.System.Collections.Generic.ListString")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 3, 2023
'@LastModified October 16, 2023

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
    names.Add DateTime.Now

    Debug.Print Strings.Format("Initial list: names.Count {0}", names.Count)
    DisplayList names
    Debug.Print
    
    names.Insert 4, "Elizabeth"
    Debug.Print Strings.Format("List after : names.Insert 4, 'Elizabeth' : names.Count {0}", names.Count)
    DisplayList names
    Debug.Print
    
    Debug.Print "names.Contains(item)"
    Debug.Print Strings.Format("Contains 'Bob' : {0}", names.Contains("Bob"))
    Debug.Print Strings.Format("Contains 'Luke' : {0}", names.Contains("Luke"))
    Debug.Print Strings.Format("Contains 'Susan' : {0}", names.Contains("Susan"))
    Debug.Print Strings.Format("Contains 'Elizabeth' : {0}", names.Contains("Elizabeth"))
    Debug.Print
    
    Dim pvtName As String
    Debug.Print "names.IndexOf(item)"
    pvtName = "James"
    Debug.Print Strings.Format("{0,3} : names.IndexOf('{1}')", names.IndexOf(pvtName), pvtName)
    pvtName = "Mark"
    Debug.Print Strings.Format("{0,3} : names.IndexOf('{1}')", names.IndexOf(pvtName), pvtName)
    pvtName = "Brian"
    Debug.Print Strings.Format("{0,3} : names.IndexOf('{1}')", names.IndexOf(pvtName), pvtName)
    pvtName = "Luke"
    Debug.Print Strings.Format("{0,3} : names.IndexOf('{1}')", names.IndexOf(pvtName), pvtName)
    Debug.Print
    
    names.Sort
    Debug.Print Strings.Format("List after : names.Sort : names.Count {0}", names.Count)
    DisplayList names
    Debug.Print
    
    'Binary search on a sorted list
    'Searches the entire sorted List<T> for an element using the default comparer and returns the zero-based index of the element.
    'The List<T> must already be sorted according to the comparer implementation; otherwise, the result is incorrect.
    Debug.Print "names.BinarySearch(item)"
    Debug.Print names.BinarySearch("James"); " : names.BinarchSearch('James')"
    Debug.Print names.BinarySearch("susan"); " : names.BinarchSearch('susan')"
    Debug.Print names.BinarySearch("Greg"); " : names.BinarchSearch('Greg')"
    Debug.Print
    
    names.Reverse
    Debug.Print Strings.Format("List after : names.Reverse : names.Count {0}", names.Count)
    DisplayList names
    Debug.Print

    names.RemoveAt 2
    Debug.Print Strings.Format("List after : names.RemoveAt 2 : names.Count {0}", names.Count)
    DisplayList names
    Debug.Print
    
    Debug.Print Strings.Format("Remove 'James' : {0}", names.Remove("James"))
    Debug.Print Strings.Format("List after : names.Remove('James') : names.Count {0}", names.Count)
    DisplayList names
    Debug.Print
    
    names.RemoveRange 2, 3
    Debug.Print Strings.Format("List after : names.RemoveRange 2, 3 : names.Count {0}", names.Count)
    DisplayList names
    Debug.Print
    
    names.Clear
    Debug.Print Strings.Format("List after : names.Clear : names.Count {0}", names.Count)
    DisplayList names
End Sub

Private Sub DisplayList(ByVal names As DotNetLib.ListString)
    Dim pvtIndex As Long
    pvtIndex = 0
    Dim varName As Variant
    For Each varName In names
        Debug.Print Strings.Format("[{0}] {1}", pvtIndex, varName)
        pvtIndex = pvtIndex + 1
    Next
End Sub

' Output:
'
'    Initial List: names.Count 7
'    [0] Bob
'    [1] Mary
'    [2] James
'    [3] Mark
'    [4] Brian
'    [5] susan
'    [6] Michael

'    List after : names.Insert 4, 'Elizabeth' : names.Count 8
'    [0] Bob
'    [1] Mary
'    [2] James
'    [3] Mark
'    [4] Elizabeth
'    [5] Brian
'    [6] susan
'    [7] Michael
'
'    names.Contains (Item)
'    Contains 'Bob' : True
'    Contains 'Luke' : False
'    Contains 'Susan' : False
'    Contains 'Elizabeth' : True
'
'    names.IndexOf (Item)
'      2 : names.IndexOf('James')
'      3 : names.IndexOf('Mark')
'      5 : names.IndexOf('Brian')
'     -1 : names.IndexOf('Luke')
'
'    List after: names.Sort: names.Count 8
'    [0] Bob
'    [1] Brian
'    [2] Elizabeth
'    [3] James
'    [4] Mark
'    [5] Mary
'    [6] Michael
'    [7] susan
'
'    names.BinarySearch (Item)
'     3  : names.BinarchSearch('James')
'     7  : names.BinarchSearch('susan')
'    -4  : names.BinarchSearch('Greg')
'
'    List after: names.Reverse: names.Count 8
'    [0] susan
'    [1] Michael
'    [2] Mary
'    [3] Mark
'    [4] James
'    [5] Elizabeth
'    [6] Brian
'    [7] Bob
'
'    List after: names.RemoveAt 2: names.Count 7
'    [0] susan
'    [1] Michael
'    [2] Mark
'    [3] James
'    [4] Elizabeth
'    [5] Brian
'    [6] Bob
'
'    Remove 'James' : True
'    List after : names.Remove('James') : names.Count 6
'    [0] susan
'    [1] Michael
'    [2] Mark
'    [3] Elizabeth
'    [4] Brian
'    [5] Bob
'
'    List after: names.RemoveRange 2, 3: names.Count 3
'    [0] susan
'    [1] Michael
'    [2] Bob
'
'    List after: names.Clear: names.Count 0


