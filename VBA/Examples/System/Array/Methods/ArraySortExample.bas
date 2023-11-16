Attribute VB_Name = "ArraySortExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 29, 2023
'@LastModified October 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.sort?view=netframework-4.8.1#system-array-sort(system-array)

'@Dependencies ReverseComparer.cls

Option Explicit

''
' The following code example shows how to sort the values in an Array using the
' default comparer and a custom comparer that reverses the sort order.
' Note that the result might vary depending on the current CultureInfo.
''
Public Sub ArraySort()
    ' Create and initialize a new array.
    Dim words As DotNetLib.Array
    Set words = Arrays.CreateInitialize1D(VBAString.GetType(), _
                        "The", "QUICK", "BROWN", "FOX", "jumps", _
                        "over", "the", "lazy", "dog")
                        
    ' Instantiate the reverse comparer.
    Dim revComparer As mscorlib.IComparer
    Set revComparer = New ReverseComparer
    
    ' Display the values of the array.
    Debug.Print "The original order of elements in the array:"
    Call DisplayValues(words)
    
    ' Sort a section of the array using the default comparer.
    Call Arrays.Sort2(words, 1, 3)
    Debug.Print "After sorting elements 1-3 by using the default comparer:"
    Call DisplayValues(words)

    ' Sort a section of the array using the reverse case-insensitive comparer.
    Call Arrays.Sort2(words, 1, 3, revComparer)
    Debug.Print "After sorting elements 1-3 by using the reverse case-insensitive comparer:"
    Call DisplayValues(words)

    ' Sort the entire array using the default comparer.
    Call Arrays.Sort(words)
    Debug.Print "After sorting the entire array by using the default comparer:"
    Call DisplayValues(words)

    ' Sort the entire array by using the reverse case-insensitive comparer.
    Call Arrays.Sort(words, revComparer)
    Debug.Print "After sorting the entire array using the reverse case-insensitive comparer:"
    Call DisplayValues(words)

End Sub

Public Sub DisplayValues(ByVal arr As DotNetLib.Array)
    Dim i As Long
    For i = arr.GetLowerBound(0) To arr.GetUpperBound(0)
        Debug.Print VBAString.Format("   [{0}] : {1}", i, arr(i))
    Next i
    Debug.Print
End Sub

' The example displays the following output:
'    The original order of elements in the array:
'       [0] : The
'       [1] : QUICK
'       [2] : BROWN
'       [3] : FOX
'       [4] : jumps
'       [5] : over
'       [6] : the
'       [7] : lazy
'       [8] : dog
'
'    After sorting elements 1-3 by using the default comparer:
'       [0] : The
'       [1] : BROWN
'       [2] : FOX
'       [3] : QUICK
'       [4] : jumps
'       [5] : over
'       [6] : the
'       [7] : lazy
'       [8] : dog
'
'    After sorting elements 1-3 by using the reverse case-insensitive comparer:
'       [0] : The
'       [1] : QUICK
'       [2] : FOX
'       [3] : BROWN
'       [4] : jumps
'       [5] : over
'       [6] : the
'       [7] : lazy
'       [8] : dog
'
'    After sorting the entire array by using the default comparer:
'       [0] : BROWN
'       [1] : dog
'       [2] : FOX
'       [3] : jumps
'       [4] : lazy
'       [5] : over
'       [6] : QUICK
'       [7] : the
'       [8] : The
'
'    After sorting the entire array using the reverse case-insensitive comparer:
'       [0] : the
'       [1] : The
'       [2] : QUICK
'       [3] : over
'       [4] : lazy
'       [5] : jumps
'       [6] : FOX
'       [7] : dog
'       [8] : BROWN


