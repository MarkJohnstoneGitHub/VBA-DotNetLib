Attribute VB_Name = "ArrayListLastIndexOfExample"
'@Folder "Examples.System.Collections.ArrayList.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 25, 2023
'@LastModified October 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.lastindexof?view=netframework-4.8.1#system-collections-arraylist-lastindexof(system-object)

Option Explicit

''
' The following code example shows how to determine the index of the last
' occurrence of a specified element.
''
Public Sub ArrayListLastIndexOf()
    ' Creates and initializes a new ArrayList with three elements of the same value.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    Call myAL.Add("the")
    Call myAL.Add("quick")
    Call myAL.Add("brown")
    Call myAL.Add("fox")
    Call myAL.Add("jumps")
    Call myAL.Add("over")
    Call myAL.Add("the")
    Call myAL.Add("lazy")
    Call myAL.Add("dog")
    Call myAL.Add("in")
    Call myAL.Add("the")
    Call myAL.Add("barn")

    ' Displays the values of the ArrayList.
    Debug.Print "The ArrayList contains the following values:"
    Call PrintIndexAndValues(myAL)

    ' Searches for the last occurrence of the duplicated value.
    Dim myString As String
    myString = "the"
    Dim myIndex As Long
    myIndex = myAL.LastIndexOf(myString)
    Debug.Print VBString.Format("The last occurrence of ""{0}"" is at index {1}.", myString, myIndex)

    ' Searches for the last occurrence of the duplicated value in the first section of the ArrayList.
     myIndex = myAL.LastIndexOf2(myString, 8)
    Debug.Print VBString.Format("The last occurrence of ""{0}"" between the start and index 8 is at index {1}.", myString, myIndex)

    ' Searches for the last occurrence of the duplicated value in a section of the ArrayList.  Note that the start index is greater than the end index because the search is done backward.
    myIndex = myAL.LastIndexOf3(myString, 10, 6)
    Debug.Print VBString.Format("The last occurrence of ""{0}"" between index 10 and index 5 is at index {1}.", myString, myIndex)
End Sub

Public Sub PrintIndexAndValues(ByVal myList As mscorlib.IEnumerable)
    Dim i As Long
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBString.Format("   [{0}]:    {1}", i, obj)
        i = i + 1
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The ArrayList contains the following values:
'   [0]:    the
'   [1]:    quick
'   [2]:    brown
'   [3]:    fox
'   [4]:    jumps
'   [5]:    over
'   [6]:    the
'   [7]:    lazy
'   [8]:    dog
'   [9]:    in
'   [10]:    the
'   [11]:    barn
'
'The last occurrence of "the" is at index 10.
'The last occurrence of "the" between the start and index 8 is at index 6.
'The last occurrence of "the" between index 10 and index 5 is at index 10.
'*/
