Attribute VB_Name = "TestArrayCreateInitialize3D"
'@Folder("Testing.System.Array")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 27, 2023
'@LastModified October 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.array.length?view=netframework-4.8.1#examples
' https://www.programiz.com/cpp-programming/multidimensional-arrays#:~:text=Initialization%20of%20three%2Ddimensional%20array&text=A%20better%20way%20to%20initialize%20this%20array%20is%3A,of%20this%20three%2Ddimensional%20array.

Option Explicit

''
' Testing creating a three-dimensional array with initial values.
''
Public Sub TestArrayCreateInitialize3D()
    Dim myArray As DotNetLib.Array
    With Arrays
        Set myArray = .CreateInitialize3D(Int32.GetType(), _
                  .AssignValues(.AssignValues(3, 4, 2, 3), .AssignValues(0, -3, 9, 11), .AssignValues(23, 12, 23, 2)), _
                  .AssignValues(.AssignValues(13, 4, 56, 3), .AssignValues(5, 9, 3, 5), .AssignValues(5, 1, 4, 9)) _
                  )
    End With
    Call ShowArrayInfo(myArray)
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
End Sub

'/*
' This code produces the following output.
'
' Length of Array:       24
' Number of Dimensions:   3
'    Dimension 1:   2
'    Dimension 2:   3
'    Dimension 3:   4
'*/
