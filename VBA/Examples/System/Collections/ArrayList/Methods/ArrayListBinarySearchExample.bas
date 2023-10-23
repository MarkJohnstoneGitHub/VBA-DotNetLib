Attribute VB_Name = "ArrayListBinarySearchExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 6, 2023
'@LastModified October 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.binarysearch?view=netframework-4.8.1#system-collections-arraylist-binarysearch(system-object)

'@Remarks BinarySearch requires searching the same data type.

Option Explicit

''
' The following code example shows how to use BinarySearch to locate a specific
' object in the ArrayList.
''
Public Sub ArrayListBinarySearch()
    ' Creates and initializes a new ArrayList. BinarySearch requires
    ' a sorted ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    Dim i As Long
    For i = 0 To 4
        myAL.Add i * 2
    Next i
    
    ' Displays the ArrayList.
    Debug.Print "The int ArrayList contains the following:"
    PrintValues myAL
    
    ' Locates a specific object that does not exist in the ArrayList.
    Dim myObjectOdd As Long
    myObjectOdd = 3
    FindMyObject myAL, myObjectOdd

    ' Locates an object that exists in the ArrayList.
    Dim myObjectEven As Long
    myObjectEven = 6
    FindMyObject myAL, myObjectEven
End Sub

Private Sub FindMyObject(ByVal myList As DotNetLib.ArrayList, ByVal myObject As Variant)
    Dim myIndex As Long
    myIndex = myList.BinarySearch(myObject)
    If (myIndex < 0) Then
        Debug.Print Strings.Format("The object to search for ({0}) is not found. The next larger object is at index {1}.", myObject, Not myIndex)
    Else
        Debug.Print Strings.Format("The object to search for ({0}) is at index {1}.", myObject, myIndex)
    End If
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print Strings.Format("   {0}", obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The int ArrayList contains the following:
'   0   2   4   6   8
'The object to search for (3) is not found. The next larger object is at index 2.
'The object to search for (6) is at index 3.
'*/
