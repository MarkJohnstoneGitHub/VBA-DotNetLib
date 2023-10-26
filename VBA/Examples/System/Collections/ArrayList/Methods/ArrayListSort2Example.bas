Attribute VB_Name = "ArrayListSort2Example"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 26, 2023
'@LastModified October 26, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.sort?view=netframework-4.8.1#system-collections-arraylist-sort(system-collections-icomparer)

'@Dependencies MyReverserClass.cls

Option Explicit

''
' The following code example shows how to sort the values in an ArrayList using
' the default comparer and a custom comparer that reverses the sort order.
''
Public Sub ArrayListSort2()
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
    Call PrintIndexAndValues(myAL)

    ' Sorts the values of the ArrayList using the default comparer.
    Call myAL.Sort
    Debug.Print "After sorting with the default comparer:"
    Call PrintIndexAndValues(myAL)
    
    ' Sorts the values of the ArrayList using the reverse case-insensitive comparer.
    Dim myComparer As mscorlib.IComparer
    Set myComparer = New MyReverserClass
    Call myAL.Sort2(myComparer)
    Debug.Print "After sorting with the reverse case-insensitive comparer:"
    Call PrintIndexAndValues(myAL)
End Sub

Public Sub PrintIndexAndValues(ByVal myList As mscorlib.IEnumerable)
    Dim formatString As String
    formatString = Regex.Unescape("\t[{0}]:\t{1}")
    Dim i As Long
    Dim obj As Variant
    For Each obj In myList
        Debug.Print Strings.Format(formatString, i, obj)
        i = i + 1
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'The ArrayList initially contains the following values:
'    [0]:    The
'    [1]:    quick
'    [2]:    brown
'    [3]:    fox
'    [4]:    jumps
'    [5]:    over
'    [6]:    the
'    [7]:    lazy
'    [8]:    dog
'
'After sorting with the default comparer:
'    [0]:    brown
'    [1]:    dog
'    [2]:    fox
'    [3]:    jumps
'    [4]:    lazy
'    [5]:    over
'    [6]:    quick
'    [7]:    the
'    [8]:    The
'
'After sorting with the reverse case-insensitive comparer:
'    [0]:    the
'    [1]:    The
'    [2]:    quick
'    [3]:    over
'    [4]:    lazy
'    [5]:    jumps
'    [6]:    fox
'    [7]:    dog
'    [8]:    brown
'*/
