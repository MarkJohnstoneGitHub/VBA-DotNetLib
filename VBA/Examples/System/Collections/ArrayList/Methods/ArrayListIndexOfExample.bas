Attribute VB_Name = "ArrayListIndexOfExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 25, 2023
'@LastModified October 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.indexof?view=netframework-4.8.1#system-collections-arraylist-indexof(system-object)

Option Explicit

''
' The following code example shows how to determine the index of the first
' occurrence of a specified element.
''
Public Sub ArrayListIndexOf()
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
    
    ' Search for the first occurrence of the duplicated value.
    Dim myString As String
    myString = "the"
    Dim myIndex As Long
    myIndex = myAL.IndexOf(myString)
    Debug.Print Strings.Format("The first occurrence of ""{0}"" is at index {1}.", myString, myIndex)
    
    ' Search for the first occurrence of the duplicated value in the last section of the ArrayList.
     myIndex = myAL.IndexOf2(myString, 4)
    Debug.Print Strings.Format("The first occurrence of ""{0}"" between index 4 and the end is at index {1}.", myString, myIndex)

    ' Search for the first occurrence of the duplicated value in a section of the ArrayList.
    myIndex = myAL.IndexOf3(myString, 6, 6)
    Debug.Print Strings.Format("The first occurrence of ""{0}"" between index 6 and index 11 is at index {1}.", myString, myIndex)
    
    ' Search for the first occurrence of the duplicated value in a small section at the end of the ArrayList.
    myIndex = myAL.IndexOf2(myString, 11)
    Debug.Print Strings.Format("The first occurrence of ""{0}"" between index 11 and the end is at index {1}.", myString, myIndex)
End Sub

Public Sub PrintIndexAndValues(ByVal myList As mscorlib.IEnumerable)
    Dim i As Long
    i = 0
    Dim obj As Variant
    For Each obj In myList
        Debug.Print Strings.Format("   [{0}]:    {1}", i, obj)
        i = i + 1
    Next
    Debug.Print
End Sub

'/*
'This code produces output similar to the following:
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
'   [10]:   the
'   [11]:   barn
'
'The first occurrence of "the" is at index 0.
'The first occurrence of "the" between index 4 and the end is at index 6.
'The first occurrence of "the" between index 6 and index 11 is at index 6.
'The first occurrence of "the" between index 11 and the end is at index -1.
'*/
