Attribute VB_Name = "ArrayListItemExample"
'@Folder "Examples.System.Collections.ArrayList.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 6, 2023
'@LastModified October 25 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.item?view=netframework-4.8.1#examples

'@Remarks
' Cannot assign value types using arraylist.Item(index) = valueType
' Eg. pvtStringList.Item(2) = "abcd" produces a Run-time error 424 Object required
' For value types added the Arraylist.SetValue(index,item) member
' Arraylist.SetValue(index,item) can be use for value or object types.
' Eg. Assigning a value type pvtStringList.SetItem 2, "abcd"

Option Explicit

Private Const Quote As String = """"

''
' The following code example creates an ArrayList and adds several items.
' The example demonstrates accessing elements with the Item[] property (the indexer in C#),
' and changing an element by assigning a new value to the Item[] property for a specified index.
' The example also shows that the Item[] property cannot be used to access or add
' elements outside the current size of the list.
''
Public Sub ArrayListItem()
    ' Create an empty ArrayList, and add some elements.
    Dim pvtStringList As DotNetLib.ArrayList
    Set pvtStringList = ArrayList.Create()
    
    pvtStringList.Add "a"
    pvtStringList.Add "abc"
    pvtStringList.Add "abcdef"
    pvtStringList.Add "abcdefg"

    ' The Item property is an indexer, so the property name is
    ' not required.
    Debug.Print VBString.Format("Element {0} is {2}{1}{2}", 2, pvtStringList(2), Quote)
    
    ' Assigning a value to the property changes the value of
    ' the indexed element
    pvtStringList.SetValue 2, "abcd" 'Note for assigning value types use SetValue(index,item) not Item(index) = item
    Debug.Print VBString.Format("Element {0} is {2}{1}{2}", 2, pvtStringList(2), Quote)
    
    ' Accessing an element outside the current element count
    ' causes an exception.
    Debug.Print VBString.Format("Number of elements in the list: {0}", _
                                pvtStringList.Count)
                                
    On Error Resume Next
    Debug.Print VBString.Format("Element {0} is {2}{1}{2}", _
                pvtStringList.Count, pvtStringList(pvtStringList.Count), Quote)
    If Catch(ArgumentOutOfRangeException) Then
        Debug.Print VBString.Format("pvtStringList({0}) is out of range.", _
                                    pvtStringList.Count)
    End If
    On Error GoTo 0 'Stop code and display error
    
    ' You cannot use the Item property to add new elements.
    On Error Resume Next
    pvtStringList.SetValue pvtStringList.Count, "42"
    If Catch(ArgumentOutOfRangeException) Then
        Debug.Print VBString.Format("pvtStringList({0}) is out of range.", _
                                    pvtStringList.Count)
    End If
    On Error GoTo 0 'Stop code and display error
    
    Debug.Print
    Dim i As Long
    For i = 0 To pvtStringList.Count - 1
        Debug.Print VBString.Format("Element {0} is {2}{1}{2}", i, _
                                    pvtStringList(i), Quote)
    Next i
    
    Debug.Print
    Dim obj As Variant
    For Each obj In pvtStringList
        Debug.Print obj
    Next
End Sub

'/*
' This code example produces the following output:
'
'Element 2 Is "abcdef"
'Element 2 Is "abcd"
'Number of elements in the list: 4
'stringList(4) is out of range.
'stringList(4) is out of range.
'
'Element 0 Is "a"
'Element 1 Is "abc"
'Element 2 Is "abcd"
'Element 3 Is "abcdefg"
'
'a
'abc
'abcd
'abcdefg
' */


