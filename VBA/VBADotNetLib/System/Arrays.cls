VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides methods for creating, manipulating, searching, and sorting arrays, thereby serving as the base class for all arrays in the common language runtime."
'@Folder "VBADotNetLib.System"
'@ModuleDescription("Provides methods for creating, manipulating, searching, and sorting arrays, thereby serving as the base class for all arrays in the common language runtime.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 8, 2023
'@LastModified October 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1

'@Remarks
' Singleton class for Array containing constructors and static members.

Option Explicit

Private Type TArraySingleton
   ArraySingleton As DotNetLib.ArraySingleton
End Type

Private this As TArraySingleton

Private Sub Class_Initialize()
   Set this.ArraySingleton = New DotNetLib.ArraySingleton
End Sub

Private Sub Class_Terminate()
   Set this.ArraySingleton = Nothing
End Sub

'@Static
'@Description("Searches an entire one-dimensional sorted array for a value using the specified IComparer interface.")
Public Function BinarySearch(ByVal pArray As DotNetLib.Array, ByVal pValue As Variant, Optional ByVal comparer As mscorlib.IComparer = Nothing) As Long
Attribute BinarySearch.VB_Description = "Searches an entire one-dimensional sorted array for a value using the specified IComparer interface."
    BinarySearch = this.ArraySingleton.BinarySearch_2(pArray, pValue, comparer)
End Function

'@Static
'@Description("Searches a range of elements in a one-dimensional sorted array for a value, using the specified IComparer interface.")
Public Function BinarySearch2(ByVal pArray As DotNetLib.Array, ByVal index As Long, ByVal pLength As Long, ByVal Value As Variant, Optional ByVal comparer As mscorlib.IComparer = Nothing) As Long
Attribute BinarySearch2.VB_Description = "Searches a range of elements in a one-dimensional sorted array for a value, using the specified IComparer interface."
   BinarySearch2 = this.ArraySingleton.BinarySearch_3(pArray, index, pLength, Value, comparer)
End Function

'@Static
'@Description("Sets a range of elements in an array to the default value of each element type.")
Public Sub Clear(ByVal pArray As DotNetLib.Array, ByVal pIndex As Long, ByVal pLength As Long)
Attribute Clear.VB_Description = "Sets a range of elements in an array to the default value of each element type."
    this.ArraySingleton.Clear pArray, pIndex, pLength
End Sub

'@Static
'@Description("Copies a range of elements from an Array starting at the specified source index and pastes them to another Array starting at the specified destination index. Guarantees that all changes are undone if the copy does not succeed completely.")
Public Sub ConstrainedCopy(ByVal sourceArray As DotNetLib.Array, ByVal sourceIndex As Long, ByVal destinationArray As DotNetLib.Array, ByVal destinationIndex As Long, ByVal pLength As Long)
Attribute ConstrainedCopy.VB_Description = "Copies a range of elements from an Array starting at the specified source index and pastes them to another Array starting at the specified destination index. Guarantees that all changes are undone if the copy does not succeed completely."
    this.ArraySingleton.ConstrainedCopy sourceArray, sourceIndex, destinationArray, destinationIndex, pLength
End Sub

'@Static
'@Description("Copies a range of elements from an Array starting at the first element and pastes them into another Array starting at the first element. The length is specified as a 32-bit integer.")
Public Sub Copy(ByVal sourceArray As DotNetLib.Array, ByVal destinationArray As DotNetLib.Array, ByVal pLength As Long)
Attribute Copy.VB_Description = "Copies a range of elements from an Array starting at the first element and pastes them into another Array starting at the first element. The length is specified as a 32-bit integer."
    this.ArraySingleton.Copy sourceArray, destinationArray, pLength
End Sub

'@Static
'@Description("Copies a range of elements from an Array starting at the first element and pastes them into another Array starting at the first element. The length is specified as a 64-bit integer.")
Public Sub Copy2(ByVal sourceArray As DotNetLib.Array, ByVal destinationArray As DotNetLib.Array, ByVal pLength As LongLong)
Attribute Copy2.VB_Description = "Copies a range of elements from an Array starting at the first element and pastes them into another Array starting at the first element. The length is specified as a 64-bit integer."
    this.ArraySingleton.Copy_2 sourceArray, destinationArray, pLength
End Sub

'@Static
'@Description("Copies a range of elements from an Array starting at the specified source index and pastes them to another Array starting at the specified destination index. The length and the indexes are specified as 32-bit integers.")
Public Sub Copy3(ByVal sourceArray As DotNetLib.Array, ByVal sourceIndex As Long, ByVal destinationArray As DotNetLib.Array, ByVal destinationIndex As Long, ByVal pLength As Long)
Attribute Copy3.VB_Description = "Copies a range of elements from an Array starting at the specified source index and pastes them to another Array starting at the specified destination index. The length and the indexes are specified as 32-bit integers."
    this.ArraySingleton.Copy_3 sourceArray, sourceIndex, destinationArray, destinationIndex, pLength
End Sub

'@Static
'@Description("Creates a one-dimensional Array of the specified Type and length, with zero-based indexing.")
Public Function CreateInstance(ByVal elementType As DotNetLib.Type, ByVal pLength As Long) As DotNetLib.Array
Attribute CreateInstance.VB_Description = "Creates a one-dimensional Array of the specified Type and length, with zero-based indexing."
   Set CreateInstance = this.ArraySingleton.CreateInstance(elementType, pLength)
End Function

'@Static
'@Description("Creates a two-dimensional Array of the specified Type and dimension lengths, with zero-based indexing.")
Public Function CreateInstance2(ByVal elementType As DotNetLib.Type, ByVal length1 As Long, ByVal length2 As Long) As DotNetLib.Array
Attribute CreateInstance2.VB_Description = "Creates a two-dimensional Array of the specified Type and dimension lengths, with zero-based indexing."
   Set CreateInstance2 = this.ArraySingleton.CreateInstance_2(elementType, length1, length2)
End Function

'@Static
'@Description("Creates a three-dimensional Array of the specified Type and dimension lengths, with zero-based indexing.")
Public Function CreateInstance3(ByVal elementType As DotNetLib.Type, ByVal length1 As Long, ByVal length2 As Long, ByVal length3 As Long) As DotNetLib.Array
Attribute CreateInstance3.VB_Description = "Creates a three-dimensional Array of the specified Type and dimension lengths, with zero-based indexing."
   Set CreateInstance3 = this.ArraySingleton.CreateInstance_3(elementType, length1, length2, length3)
End Function

'@Static
'@Description("Creates a multidimensional Array of the specified Type and dimension lengths, with zero-based indexing. The dimension lengths are specified in an array of 32-bit integers.")
Public Function CreateInstance4(ByVal elementType As DotNetLib.Type, ByRef lengths() As Long) As DotNetLib.Array
Attribute CreateInstance4.VB_Description = "Creates a multidimensional Array of the specified Type and dimension lengths, with zero-based indexing. The dimension lengths are specified in an array of 32-bit integers."
   Set CreateInstance4 = this.ArraySingleton.CreateInstance_4(elementType, lengths)
End Function

'@Static
'@Description("Creates a multidimensional Array of the specified Type and dimension lengths, with zero-based indexing. The dimension lengths are specified in an array of 64-bit integers.")
Public Function CreateInstance5(ByVal elementType As DotNetLib.Type, ByRef lengths() As LongLong) As DotNetLib.Array
Attribute CreateInstance5.VB_Description = "Creates a multidimensional Array of the specified Type and dimension lengths, with zero-based indexing. The dimension lengths are specified in an array of 64-bit integers."
   Set CreateInstance5 = this.ArraySingleton.CreateInstance_5(elementType, lengths)
End Function

'@Static
'@Description("Creates a multidimensional Array of the specified Type and dimension lengths, with the specified lower bounds.")
Public Function CreateInstance6(ByVal elementType As DotNetLib.Type, ByRef lengths() As Long, ByRef lowerBounds() As Long) As DotNetLib.Array
Attribute CreateInstance6.VB_Description = "Creates a multidimensional Array of the specified Type and dimension lengths, with the specified lower bounds."
   Set CreateInstance6 = this.ArraySingleton.CreateInstance_6(elementType, lengths, lowerBounds)
End Function

'@Static
'@Description("Searches for the specified object and returns the index of its first occurrence in a one-dimensional array.")
Public Function IndexOf(ByVal pArray As DotNetLib.Array, ByVal Value As Variant) As Long
Attribute IndexOf.VB_Description = "Searches for the specified object and returns the index of its first occurrence in a one-dimensional array."
   IndexOf = this.ArraySingleton.IndexOf(pArray, Value)
End Function

'@Static
'@Description("Searches for the specified object in a range of elements of a one-dimensional array, and returns the index of its first occurrence. The range extends from a specified index to the end of the array.")
Public Function IndexOf2(ByVal pArray As DotNetLib.Array, ByVal Value As Variant, ByVal startIndex As Long) As Long
Attribute IndexOf2.VB_Description = "Searches for the specified object in a range of elements of a one-dimensional array, and returns the index of its first occurrence. The range extends from a specified index to the end of the array."
   IndexOf2 = this.ArraySingleton.IndexOf_2(pArray, Value, startIndex)
End Function

'@Static
'@Description("Searches for the specified object in a range of elements of a one-dimensional array, and returns the index of ifs first occurrence. The range extends from a specified index for a specified number of elements.")
Public Function IndexOf3(ByVal pArray As DotNetLib.Array, ByVal Value As Variant, ByVal startIndex As Long, ByVal pCount As Long) As Long
Attribute IndexOf3.VB_Description = "Searches for the specified object in a range of elements of a one-dimensional array, and returns the index of ifs first occurrence. The range extends from a specified index for a specified number of elements."
   IndexOf3 = this.ArraySingleton.IndexOf_3(pArray, Value, startIndex, pCount)
End Function

'@Static
'@Description("Searches for the specified object and returns the index of the last occurrence within the entire one-dimensional Array.")
Public Function LastIndexOf(ByVal pArray As DotNetLib.Array, ByVal Value As Variant) As Long
Attribute LastIndexOf.VB_Description = "Searches for the specified object and returns the index of the last occurrence within the entire one-dimensional Array."
   LastIndexOf = this.ArraySingleton.LastIndexOf(pArray, Value)
End Function

'@Static
'@Description("Searches for the specified object and returns the index of the last occurrence within the range of elements in the one-dimensional Array that extends from the first element to the specified index.")
Public Function LastIndexOf2(ByVal pArray As DotNetLib.Array, ByVal Value As Variant, ByVal startIndex As Long) As Long
Attribute LastIndexOf2.VB_Description = "Searches for the specified object and returns the index of the last occurrence within the range of elements in the one-dimensional Array that extends from the first element to the specified index."
   LastIndexOf2 = this.ArraySingleton.LastIndexOf_2(pArray, Value, startIndex)
End Function

'@Static
'@Description("Searches for the specified object and returns the index of the last occurrence within the range of elements in the one-dimensional Array that contains the specified number of elements and ends at the specified index.")
Public Function LastIndexOf3(ByVal pArray As DotNetLib.Array, ByVal Value As Variant, ByVal startIndex As Long, ByVal pCount As Long) As Long
Attribute LastIndexOf3.VB_Description = "Searches for the specified object and returns the index of the last occurrence within the range of elements in the one-dimensional Array that contains the specified number of elements and ends at the specified index."
   LastIndexOf3 = this.ArraySingleton.LastIndexOf_3(pArray, Value, startIndex, pCount)
End Function

'@Static
'@Description("Changes the number of elements of a one-dimensional array to the specified new size.")
Public Sub Resize(ByRef pArray As DotNetLib.Array, ByVal newSize As Long)
Attribute Resize.VB_Description = "Changes the number of elements of a one-dimensional array to the specified new size."
    this.ArraySingleton.Resize pArray, newSize
End Sub

'@Static
'@Description("Reverses the sequence of the elements in the entire one-dimensional Array.")
Public Sub Reverse(ByVal pArray As DotNetLib.Array)
Attribute Reverse.VB_Description = "Reverses the sequence of the elements in the entire one-dimensional Array."
    this.ArraySingleton.Reverse pArray
End Sub

'@Static
'@Description("Reverses the sequence of a subset of the elements in the one-dimensional Array.")
Public Sub Reverse2(ByVal pArray As DotNetLib.Array, ByVal index As Long, ByVal pLength As Long)
Attribute Reverse2.VB_Description = "Reverses the sequence of a subset of the elements in the one-dimensional Array."
    this.ArraySingleton.Reverse_2 pArray, index, pLength
End Sub

'Static
'@Description("Sorts the elements in an entire one-dimensional Array using the IComparable implementation of each element of the Array.")
Public Sub Sort(ByVal pArray As DotNetLib.Array, Optional ByVal comparer As mscorlib.IComparer = Nothing)
Attribute Sort.VB_Description = "Sorts the elements in an entire one-dimensional Array using the IComparable implementation of each element of the Array."
    this.ArraySingleton.Sort pArray, comparer
End Sub

'Static
'@Description("Sorts the elements in a range of elements in a one-dimensional Array using the specified IComparer.")
Public Sub Sort2(ByVal pArray As DotNetLib.Array, ByVal index As Long, ByVal pLength As Long, Optional ByVal comparer As mscorlib.IComparer = Nothing)
Attribute Sort2.VB_Description = "Sorts the elements in a range of elements in a one-dimensional Array using the specified IComparer."
    this.ArraySingleton.Sort_2 pArray, index, pLength, comparer
End Sub

'Static
'@Description("Sorts a pair of one-dimensional Array objects (one contains the keys and the other contains the corresponding items) based on the keys in the first Array using the specified IComparer.")
Public Sub Sort3(ByVal keys As DotNetLib.Array, ByVal items As DotNetLib.Array, Optional ByVal comparer As mscorlib.IComparer = Nothing)
Attribute Sort3.VB_Description = "Sorts a pair of one-dimensional Array objects (one contains the keys and the other contains the corresponding items) based on the keys in the first Array using the specified IComparer."
    this.ArraySingleton.Sort_3 keys, items, comparer
End Sub

'Static
'@Description("Sorts a range of elements in a pair of one-dimensional Array objects (one contains the keys and the other contains the corresponding items) based on the keys in the first Array using the specified IComparer.")
Public Sub Sort4(ByVal keys As DotNetLib.Array, ByVal items As DotNetLib.Array, ByVal index As Long, ByVal pLength As Long, Optional ByVal comparer As mscorlib.IComparer = Nothing)
Attribute Sort4.VB_Description = "Sorts a range of elements in a pair of one-dimensional Array objects (one contains the keys and the other contains the corresponding items) based on the keys in the first Array using the specified IComparer."
    this.ArraySingleton.Sort_4 keys, items, index, pLength, comparer
End Sub

'@TODO FindAll
'Public Function FindAll(ByVal values As DotNetLib.Array, pMatch) As DotNetLib.Array
'End Function

'--------------------------------------------------------------------
'Extensions
'Added to create and initialize array values
'--------------------------------------------------------------------

'Static
'@Description("Creates a one-dimensional Array of the specified Type with initial values")
Public Function CreateInitialize1D(ByVal elementType As DotNetLib.Type, ParamArray Values() As Variant) As DotNetLib.Array
Attribute CreateInitialize1D.VB_Description = "Creates a one-dimensional Array of the specified Type with initial values"
    Dim outArray As DotNetLib.Array
'    On Error GoTo CatchError
    'Create Array of size off parameter array
    Set outArray = Me.CreateInstance(elementType, UBound(Values) + 1)
    Dim varValue As Variant
    Dim pvtIndex As Long
    pvtIndex = 0
    For Each varValue In Values
        Call outArray.SetValue(varValue, pvtIndex)
        pvtIndex = pvtIndex + 1
    Next
    Set CreateInitialize1D = outArray
'Exit Function
'CatchError:
'    Err.Raise Number:=Err.Number, Source:="Arrays.CreateInstanceInitialize", Description:=Strings.Format("Object {0} cannot be stored in an array of type {1}.", varValue, elementType.FullName)
End Function

'Static
'@Description("Creates a two-dimensional Array of the specified Type with initial values")
Public Function CreateInitialize2D(ByVal elementType As DotNetLib.Type, ParamArray rows() As Variant) As DotNetLib.Array
Attribute CreateInitialize2D.VB_Description = "Creates a two-dimensional Array of the specified Type with initial values"
    Dim numOfRows As Long
    Dim numOfCols As Long
    numOfRows = UBound(rows) + 1
    'Two approaches either find the max columns, or first parameter defines max columns?
    Dim row As Variant
    For Each row In rows
        If UBound(row) + 1 > numOfCols Then
            numOfCols = UBound(row) + 1
        End If
    Next
    'Create Array
    Dim outArray As DotNetLib.Array
    Set outArray = Me.CreateInstance2(elementType, numOfRows, numOfCols)
    
    'Assign initial values
    Dim rowIndex As Long
    rowIndex = 0
    For Each row In rows
        Dim columnIndex As Long
        columnIndex = 0
        Dim varValue As Variant
        For Each varValue In row
            Call outArray.SetValue_2(varValue, rowIndex, columnIndex)
            columnIndex = columnIndex + 1
        Next
        rowIndex = rowIndex + 1
    Next
    Set CreateInitialize2D = outArray
End Function

'Static
'@Description("Creates a three-dimensional Array of the specified Type with initial values")
Public Function CreateInitialize3D(ByVal elementType As DotNetLib.Type, ParamArray pValues() As Variant) As DotNetLib.Array
Attribute CreateInitialize3D.VB_Description = "Creates a three-dimensional Array of the specified Type with initial values"
    Dim widthSize As Long
    Dim heightSize As Long
    Dim depthSize As Long
    widthSize = UBound(pValues) + 1
    Dim height As Variant
    For Each height In pValues
        'Get max heightSize
        If UBound(height) + 1 > heightSize Then
            heightSize = UBound(height) + 1
        End If
        Dim depth As Variant
        For Each depth In height
            If UBound(depth) + 1 > depthSize Then
                depthSize = UBound(depth) + 1
            End If
        Next
    Next
    
    'Create 3D Array
    Dim outArray As DotNetLib.Array
    Set outArray = Me.CreateInstance3(elementType, widthSize, heightSize, depthSize)

    'Assign initial values
    Dim widthIndex As Long
    widthIndex = 0
    Dim width As Variant
    For Each width In pValues
        Dim heightIndex As Long
        heightIndex = 0
        For Each height In width
            Dim depthIndex As Long
            depthIndex = 0
            Dim varValue As Variant
            For Each varValue In height
                Call outArray.SetValue_3(varValue, widthIndex, heightIndex, depthIndex)
                depthIndex = depthIndex + 1
            Next
            heightIndex = heightIndex + 1
        Next
        widthIndex = widthIndex + 1
    Next
    Set CreateInitialize3D = outArray
End Function
