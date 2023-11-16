Attribute VB_Name = "ArrayClearExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 28, 2023
'@LastModified October 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.clear?view=netframework-4.8.1#system-array-clear(system-array-system-int32-system-int32)

Option Explicit

''
' The following example uses the Clear method to reset integer values in a
' one-dimensional, two-dimensional, and three-dimensional array.
''
Public Sub ArrayClearExample()
    Debug.Print "One dimension (Rank=1):"
    Dim numbers1 As DotNetLib.Array
    Set numbers1 = Arrays.CreateInitialize1D(Int32.GetType(), 1, 2, 3, 4, 5, 6, 7, 8, 9)
    
    Dim i As Long
    For i = 0 To 8
        Debug.Print VBAString.Format("{0} ", numbers1(i));
    Next
    Debug.Print
    Debug.Print
    
    Debug.Print "Array.Clear(numbers1, 2, 5)"
    Call Arrays.Clear(numbers1, 2, 5)
    
    For i = 0 To 8
        Debug.Print VBAString.Format("{0} ", numbers1(i));
    Next
    Debug.Print
    Debug.Print
    
    Debug.Print "Two dimensions (Rank=2):"
    Dim numbers2 As DotNetLib.Array
    Set numbers2 = Arrays.CreateInitialize2D(Int32.GetType(), _
                    Array(1, 2, 3), _
                    Array(4, 5, 6), _
                    Array(7, 8, 9) _
                    )
    Dim j As Long
    For i = 0 To 2
        For j = 0 To 2
            Debug.Print VBAString.Format("{0} ", numbers2.GetValue_2(i, j));
        Next j
        Debug.Print
    Next i

    Debug.Print
    Debug.Print "Array.Clear(numbers2, 2, 5)"
    Call Arrays.Clear(numbers2, 2, 5)
    
    For i = 0 To 2
        For j = 0 To 2
            Debug.Print VBAString.Format("{0} ", numbers2.GetValue_2(i, j));
        Next j
        Debug.Print
    Next i
    
    
    Debug.Print "Three dimensions (Rank=3):"
    Dim numbers3 As DotNetLib.Array
    Set numbers3 = Arrays.CreateInitialize3D(Int32.GetType(), _
                    Array(Array(1, 2), Array(3, 4)), _
                    Array(Array(5, 6), Array(7, 8)), _
                    Array(Array(9, 10), Array(11, 12)) _
                  )
                  
    Dim k As Long
    For i = 0 To 2
        For j = 0 To 1
            For k = 0 To 1
                Debug.Print VBAString.Format("{0} ", numbers3.GetValue_3(i, j, k));
            Next k
            Debug.Print
        Next j
        Debug.Print
    Next i
    
    Debug.Print "Array.Clear(numbers3, 2, 5)"
    Call Arrays.Clear(numbers3, 2, 5)

    For i = 0 To 2
        For j = 0 To 1
            For k = 0 To 1
                Debug.Print VBAString.Format("{0} ", numbers3.GetValue_3(i, j, k));
            Next k
            Debug.Print
        Next j
        Debug.Print
    Next i
End Sub

'/*  This code example produces the following output:
' *
' * One dimension (Rank=1):
' * 1 2 3 4 5 6 7 8 9
' *
' * Array.Clear(numbers1, 2, 5)
' * 1 2 0 0 0 0 0 8 9
' *
' * Two dimensions (Rank=2):
' * 1 2 3
' * 4 5 6
' * 7 8 9
' *
' * Array.Clear(numbers2, 2, 5)
' * 1 2 0
' * 0 0 0
' * 0 8 9
' *
' * Three dimensions (Rank=3):
' * 1 2
' * 3 4
' *
' * 5 6
' * 7 8
' *
' * 9 10
' * 11 12
' *
' * Array.Clear(numbers3, 2, 5)
' * 1 2
' * 0 0
' *
' * 0 0
' * 0 8
' *
' * 9 10
' * 11 12
' */


