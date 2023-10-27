Attribute VB_Name = "ArrayLengthExample"
'@Folder("Examples.System.Array.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 27, 2023
'@LastModified October 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.length?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the Length property to get the total number of
' elements in an array. It also uses the GetUpperBound method to determine
' the number of elements in each dimension of a multidimensional array.
''
Public Sub ArrayLength()
    ' Declare a single-dimensional string array
    Dim array1d As DotNetLib.Array
    Set array1d = Arrays.CreateInitialize1D(Strings.GetType(), "zero", "one", "two", "three")
    Call ShowArrayInfo(array1d)
    
    ' Declare a two-dimensional string array
    Dim array2d As DotNetLib.Array
    With Arrays
        Set array2d = .CreateInitialize2D(Strings.GetType(), _
                    .AssignValues("zero", "0"), _
                    .AssignValues("one", "1"), _
                    .AssignValues("two", "2"), _
                    .AssignValues("three", "3"), _
                    .AssignValues("four", "4"), _
                    .AssignValues("five", "5") _
                    )
    End With
    Call ShowArrayInfo(array2d)

    ' Declare a three-dimensional integer array
    Dim array3d As DotNetLib.Array
    With Arrays
        Set array3d = .CreateInitialize3D(Int32.GetType(), _
                  .AssignValues(.AssignValues(1, 2, 3), .AssignValues(4, 5, 6)), _
                  .AssignValues(.AssignValues(7, 8, 9), .AssignValues(10, 11, 12)) _
                  )
    End With
    Call ShowArrayInfo(array3d)
End Sub

Private Sub ShowArrayInfo(ByVal arr As DotNetLib.Array)
    Debug.Print Strings.Format("Length of Array:      {0,3}", arr.Length)
    Debug.Print Strings.Format("Number of Dimensions: {0,3}", arr.Rank)
    ' For multidimensional arrays, show number of elements in each dimension.
    If (arr.Rank > 1) Then
        Dim dimension As Long
        For dimension = 1 To arr.Rank
            Debug.Print Strings.Format("   Dimension {0}: {1,3}", dimension, _
                                        arr.GetUpperBound(dimension - 1) + 1)
        Next dimension
    End If
    Debug.Print
End Sub

' The example displays the following output:
'       Length of Array:        4
'       Number of Dimensions:   1
'
'       Length of Array:       12
'       Number of Dimensions:   2
'          Dimension 1:   6
'          Dimension 2:   2
'
'       Length of Array:       12
'       Number of Dimensions:   3
'          Dimension 1:   2
'          Dimension 2:   2
'          Dimension 3:   3
