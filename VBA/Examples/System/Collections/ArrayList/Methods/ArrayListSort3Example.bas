Attribute VB_Name = "ArrayListSort3Example"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 26, 2023
'@LastModified October 26, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.sort?view=netframework-4.8.1#system-collections-arraylist-sort(system-int32-system-int32-system-collections-icomparer)

'@Dependencies MyReverserClass.cls

Option Explicit

''
' The following code example shows how to sort the values in a range of
' elements in an ArrayList using the default comparer and a custom comparer
' that reverses the sort order.
''
Public Sub ArrayListSort3()
    ' Creates and initializes a new ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    Call myAL.Add("The")
    Call myAL.Add("QUICK")
    Call myAL.Add("BROWN")
    Call myAL.Add("FOX")
    Call myAL.Add("jumps")
    Call myAL.Add("over")
    Call myAL.Add("the")
    Call myAL.Add("lazy")
    Call myAL.Add("dog")
    
    ' Displays the values of the ArrayList.
    Debug.Print "The ArrayList initially contains the following values:"
    Call PrintIndexAndValues(myAL)

    ' Sorts the values of the ArrayList using the default comparer.
    Call myAL.Sort3(1, 3, Nothing)
    Debug.Print "After sorting from index 1 to index 3 with the default comparer:"
    Call PrintIndexAndValues(myAL)
    
    ' Sorts the values of the ArrayList using the reverse case-insensitive comparer.
    Dim MyComparer As mscorlib.IComparer
    Set MyComparer = New MyReverserClass
    Call myAL.Sort3(1, 3, MyComparer)
    Debug.Print "After sorting from index 1 to index 3 with the reverse case-insensitive comparer:"
    Call PrintIndexAndValues(myAL)
End Sub

Public Sub PrintIndexAndValues(ByVal myList As mscorlib.IEnumerable)
    Dim formatString As String
    formatString = Regex.Unescape("\t[{0}]:\t{1}")
    Dim i As Long
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBAString.Format(formatString, i, obj)
        i = i + 1
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'The ArrayList initially contains the following values:
'    [0]:    The
'    [1]:    QUICK
'    [2]:    BROWN
'    [3]:    FOX
'    [4]:    jumps
'    [5]:    over
'    [6]:    the
'    [7]:    lazy
'    [8]:    dog
'
'After sorting from index 1 to index 3 with the default comparer:
'    [0]:    The
'    [1]:    BROWN
'    [2]:    FOX
'    [3]:    QUICK
'    [4]:    jumps
'    [5]:    over
'    [6]:    the
'    [7]:    lazy
'    [8]:    dog
'
'After sorting from index 1 to index 3 with the reverse case-insensitive comparer:
'    [0]:    The
'    [1]:    QUICK
'    [2]:    FOX
'    [3]:    BROWN
'    [4]:    jumps
'    [5]:    over
'    [6]:    the
'    [7]:    lazy
'    [8]:    dog
'*/

