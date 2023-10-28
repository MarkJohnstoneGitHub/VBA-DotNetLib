Attribute VB_Name = "ArrayGetLowerBoundExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 28, 2023
'@LastModified October 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.getlowerbound?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the GetLowerBound and GetUpperBound methods to
' display the bounds of a one-dimensional and two-dimensional array and to
' display the values of their array elements.
''
Public Sub ArrayGetLowerBound()
    Dim integers As DotNetLib.Array
    Set integers = Arrays.CreateInitialize1D(Int32.GetType(), 2, 4, 6, 8, 10, 12, 14, 16, 18, 20)
    ' Get the upper and lower bound of the array.
    Dim upper As Long
    upper = integers.GetUpperBound(0)
    Dim lower As Long
    lower = integers.GetLowerBound(0)
    
    Debug.Print Strings.Format("Elements from index {0} to {1}:", lower, upper)
    ' Iterate the array.
    Dim ctr As Long
    For ctr = lower To upper
        Debug.Print Strings.Format("{0}{1}{2}", IIf(ctr = lower, "   ", ""), integers(ctr), IIf(ctr < upper, ", ", VBA.vbNewLine));
    Next
    Debug.Print

    Dim integers2d As DotNetLib.Array
    Set integers2d = Arrays.CreateInitialize2D(Int32.GetType(), _
                    Array(2, 4), Array(3, 9), Array(4, 16), Array(5, 25), _
                    Array(6, 36), Array(7, 49), Array(8, 64), Array(9, 81) _
                    )
    ' Get the number of dimensions.
    Dim pvtRank As Long
    pvtRank = integers2d.Rank
    Debug.Print Strings.Format("Number of dimensions: {0}", pvtRank)
    For ctr = 0 To pvtRank - 1
        Debug.Print Strings.Format("   Dimension {0}: ", ctr) & _
                    Strings.Format("from {0} to {1}", integers2d.GetLowerBound(ctr), integers2d.GetUpperBound(ctr))
    Next ctr

    ' Iterate the 2-dimensional array and display its values.
    Dim formatString As String
    formatString = "      {{{0}, {1}}} = {2}"
    Debug.Print "   Values of array elements:"
    Dim outer As Long
    For outer = integers2d.GetLowerBound(0) To integers2d.GetUpperBound(0)
        Dim inner As Long
        For inner = integers2d.GetLowerBound(1) To integers2d.GetUpperBound(1)
            Debug.Print Strings.Format(formatString, outer, inner, integers2d.GetValue_2(outer, inner))
        Next inner
    Next outer
End Sub

' The example displays the following output:
'       Elements from index 0 to 9:
'          2, 4, 6, 8, 10, 12, 14, 16, 18, 20
'
'       Number of dimensions: 2
'          Dimension 0: from 0 to 7
'          Dimension 1: from 0 to 1
'          Values of array elements:
'             {0, 0} = 2
'             {0, 1} = 4
'             {1, 0} = 3
'             {1, 1} = 9
'             {2, 0} = 4
'             {2, 1} = 16
'             {3, 0} = 5
'             {3, 1} = 25
'             {4, 0} = 6
'             {4, 1} = 36
'             {5, 0} = 7
'             {5, 1} = 49
'             {6, 0} = 8
'             {6, 1} = 64
'             {7, 0} = 9
'             {7, 1} = 81
