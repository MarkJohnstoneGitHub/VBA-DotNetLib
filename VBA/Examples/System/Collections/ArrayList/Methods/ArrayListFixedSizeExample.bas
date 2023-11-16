Attribute VB_Name = "ArrayListFixedSizeExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 24, 2023
'@LastModified October 24, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.fixedsize?view=netframework-4.8.1#system-collections-arraylist-fixedsize(system-collections-arraylist)

Option Explicit

''
' The following code example shows how to create a fixed-size wrapper around
' an ArrayList.
''
Public Sub ArrayListFixedSize()
    ' Creates and initializes a new ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    Call myAL.Add("The")
    Call myAL.Add("quick")
    Call myAL.Add("brown")
    Call myAL.Add("fox")
    Call myAL.Add("jumps")
    Call myAL.Add("over")
    Call myAL.Add("the")
    Call myAL.Add("lazy")
    Call myAL.Add("dog")
    
    ' Create a fixed-size wrapper around the ArrayList.
    Dim myFixedSizeAL As DotNetLib.ArrayList
    Set myFixedSizeAL = ArrayList.FixedSize(myAL)
    
    ' Display whether the ArrayLists have a fixed size or not.
    Debug.Print VBAString.Format("myAL {0}.", IIf(myAL.IsFixedSize, "has a fixed size", "does not have a fixed size"))
    Debug.Print VBAString.Format("myFixedSizeAL {0}.", IIf(myFixedSizeAL.IsFixedSize, "has a fixed size", "does not have a fixed size"))
    Debug.Print
    
    ' Display both ArrayLists.
    Debug.Print "Initially,"
    Debug.Print "Standard  :";
    Call PrintValues(myAL, " ")
    Debug.Print "Fixed size:";
    Call PrintValues(myFixedSizeAL, " ")
    
    ' Sort is allowed in the fixed-size ArrayList.
    Call myFixedSizeAL.Sort
    
    '/ Display both ArrayLists.
    Debug.Print "After Sort,"
    Debug.Print "Standard  :";
    Call PrintValues(myAL, " ")
    Debug.Print "Fixed size:";
    Call PrintValues(myFixedSizeAL, " ")
    
    ' Reverse is allowed in the fixed-size ArrayList.
    Call myFixedSizeAL.Reverse

    ' Display both ArrayLists.
    Debug.Print "After Reverse,"
    Debug.Print "Standard  :";
    Call PrintValues(myAL, " ")
    Debug.Print "Fixed size:";
    Call PrintValues(myFixedSizeAL, " ")

    ' Add an element to the standard ArrayList.
    Call myAL.Add("AddMe")

    ' Display both ArrayLists.
    Debug.Print "After adding to the standard ArrayList,"
    Debug.Print "Standard  :";
    Call PrintValues(myAL, " ")
    Debug.Print "Fixed size:";
    Call PrintValues(myFixedSizeAL, " ")
    Debug.Print

    ' Adding or inserting elements to the fixed-size ArrayList throws an exception.
    On Error Resume Next
    Call myFixedSizeAL.Add("AddMe2")
    If Err.number Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    On Error Resume Next
    Call myFixedSizeAL.Insert(3, "InsertMe")
    If Err.number Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable, ByVal mySeparator As String)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBAString.Format("{0}{1}", mySeparator, obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'myAL does not have a fixed size.
'myFixedSizeAL has a fixed size.
'
'Initially,
'Standard  : The quick brown fox jumps over the lazy dog
'Fixed size: The quick brown fox jumps over the lazy dog
'After Sort,
'Standard  : brown dog fox jumps lazy over quick the The
'Fixed size: brown dog fox jumps lazy over quick the The
'After Reverse,
'Standard  : The the quick over lazy jumps fox dog brown
'Fixed size: The the quick over lazy jumps fox dog brown
'After adding to the standard ArrayList,
'Standard  : The the quick over lazy jumps fox dog brown AddMe
'Fixed size: The the quick over lazy jumps fox dog brown AddMe
'
'Exception: System.NotSupportedException: Collection was of a fixed size.
'   at System.Collections.FixedSizeArrayList.Add(Object obj)
'   at SamplesArrayList.Main()
'Exception: System.NotSupportedException: Collection was of a fixed size.
'   at System.Collections.FixedSizeArrayList.Insert(int index, Object obj)
'   at SamplesArrayList.Main()
'
'*/
