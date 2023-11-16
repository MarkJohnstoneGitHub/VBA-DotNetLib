Attribute VB_Name = "TestArrayCreateInitialize2D"
'@Folder("Testing.System.Array")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 27, 2023
'@LastModified October 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.length?view=netframework-4.8.1#examples

Option Explicit

''
' Testing creating a two-dimensional array with initial values.
''
Public Sub TestCreateInitialize2D()
    Dim myArray As DotNetLib.Array
    With Arrays
        Set myArray = .CreateInitialize2D(VBAString.GetType(), _
                    Array("zero", "0"), _
                    Array("one", "1"), _
                    Array("two", "2"), _
                    Array("three", "3"), _
                    Array("four", "4"), _
                    Array("five", "5") _
                    )
    End With
    Call ShowArrayInfo(myArray)
End Sub

Private Sub ShowArrayInfo(ByVal arr As DotNetLib.Array)
    Debug.Print VBAString.Format("Length of Array:      {0,3}", arr.length)
    Debug.Print VBAString.Format("Number of Dimensions: {0,3}", arr.Rank)
    ' For multidimensional arrays, show number of elements in each dimension.
    If (arr.Rank > 1) Then
        Dim dimension As Long
        For dimension = 1 To arr.Rank
            Debug.Print VBAString.Format("   Dimension {0}: {1,3}", dimension, _
                                        arr.GetUpperBound(dimension - 1) + 1)
        Next dimension
    End If
End Sub

'/*
' This code produces the following output.
'
' Length of Array:       12
' Number of Dimensions:   2
'    Dimension 1:   6
'    Dimension 2:   2
'*/


