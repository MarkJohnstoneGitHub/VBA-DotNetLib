Attribute VB_Name = "ArrayListIsFixedSizeExample"
'@Folder("Examples.System.Collections.ArrayList.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 6, 2023
'@LastModified October 6, 2023

'Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.isfixedsize?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to create a fixed-size wrapper around an ArrayList.
''
Public Sub ArrayListIsFixedSize()
    ' Creates and initializes a new ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    myAL.Add "The"
    myAL.Add "quick"
    myAL.Add "brown"
    myAL.Add "fox"
    myAL.Add "jumps"
    myAL.Add "over"
    myAL.Add "the"
    myAL.Add "lazy"
    myAL.Add "dog"
    
    ' Create a fixed-size wrapper around the ArrayList.
    Dim myFixedSizeAL As DotNetLib.ArrayList
    Set myFixedSizeAL = ArrayList.FixedSize(myAL)

    ' Display whether the ArrayLists have a fixed size or not.
    Debug.Print Strings.Format("myAL {0}.", IIf(myAL.IsFixedSize, "has a fixed size", "does not have a fixed size"))
    Debug.Print Strings.Format("myFixedSizeAL {0}.", IIf(myFixedSizeAL.IsFixedSize, "has a fixed size", "does not have a fixed size"))
    Debug.Print
    
    ' Display both ArrayLists.
    Debug.Print "Initially,"
    Debug.Print "Standard  :";
    PrintValues myAL, " "
    Debug.Print "Fixed size:";
    PrintValues myFixedSizeAL, " "
    
    ' Sort is allowed in the fixed-size ArrayList.
    myFixedSizeAL.Sort
    
    ' Display both ArrayLists.
    Debug.Print "After Sort,"
    Debug.Print "Standard  :";
    PrintValues myAL, " "
    Debug.Print "Fixed size:";
    PrintValues myFixedSizeAL, " "
    
    ' Reverse is allowed in the fixed-size ArrayList.
    myFixedSizeAL.Reverse
    
    ' Display both ArrayLists.
    Debug.Print "After Reverse,"
    Debug.Print "Standard  :";
    PrintValues myAL, " "
    Debug.Print "Fixed size:";
    PrintValues myFixedSizeAL, " "
    
    ' Add an element to the standard ArrayList.
    myAL.Add "AddMe"
    
    ' Display both ArrayLists.
    Debug.Print "After adding to the standard ArrayList,"
    Debug.Print "Standard  :";
    PrintValues myAL, " "
    Debug.Print "Fixed size:";
    PrintValues myFixedSizeAL, " "
    Debug.Print
    
    ' Adding or inserting elements to the fixed-size ArrayList throws an exception.
    On Error Resume Next
    myFixedSizeAL.Add "AddMe2"
    If Catch() Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    
    On Error Resume Next
    myFixedSizeAL.Insert 3, "InsertMe"
    If Catch() Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable, ByVal mySeparator As String)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print Strings.Format("{0}{1}", mySeparator, obj);
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
'Exception: Collection was of a fixed size.
'Exception: Collection was of a fixed size.
'
'*/
