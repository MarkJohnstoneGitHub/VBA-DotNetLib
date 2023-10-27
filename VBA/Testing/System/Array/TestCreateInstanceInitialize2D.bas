Attribute VB_Name = "TestCreateInstanceInitialize2D"
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
Public Sub TestCreateInstanceInitialize2D()
    Dim myArray As DotNetLib.Array
    With Arrays
        Set myArray = .CreateInstanceInitialize2D(Strings.GetType(), _
                    .InitializeRow("zero", "0"), _
                    .InitializeRow("one", "1"), _
                    .InitializeRow("two", "2"), _
                    .InitializeRow("three", "3"), _
                    .InitializeRow("four", "4"), _
                    .InitializeRow("five", "5") _
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


