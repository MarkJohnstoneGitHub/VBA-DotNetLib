Attribute VB_Name = "ArrayListSortExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 25, 2023
'@LastModified October 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.sort?view=netframework-4.8.1#system-collections-arraylist-sort

Option Explicit

''
' The following code example shows how to sort the values in an ArrayList.
''
Public Sub ArrayListSort()
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
    
    ' Displays the values of the ArrayList.
    Debug.Print "The ArrayList initially contains the following values:"
    Call PrintValues(myAL)

    ' Sorts the values of the ArrayList.
    Call myAL.Sort

    ' Displays the values of the ArrayList.
    Debug.Print "After sorting:"
    Call PrintValues(myAL)
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print Strings.Format("   {0}", obj)
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The ArrayList initially contains the following values:
'   The
'   quick
'   brown
'   fox
'   jumps
'   over
'   the
'   lazy
'   dog
'
'After sorting:
'   brown
'   dog
'   fox
'   jumps
'   lazy
'   over
'   quick
'   the
'   The
'*/
