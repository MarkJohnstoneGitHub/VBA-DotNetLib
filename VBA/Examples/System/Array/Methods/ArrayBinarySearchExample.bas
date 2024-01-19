Attribute VB_Name = "ArrayBinarySearchExample"
'@Folder "Examples.System.Array.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 28, 2023
'@LastModified October 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.binarysearch?view=netframework-4.8.1#system-array-binarysearch(system-array-system-object)

'@Remarks
' Define the object searching for as the same data type of the Array.
Option Explicit

Public Sub ArrayBinarySearch()
    ' Creates and initializes a new Array.
    Dim myIntArray As DotNetLib.Array
    Set myIntArray = Arrays.CreateInstance(Int32.GetType(), 5)
    Call myIntArray.SetValue(8, 0)
    Call myIntArray.SetValue(2, 1)
    Call myIntArray.SetValue(6, 2)
    Call myIntArray.SetValue(3, 3)
    Call myIntArray.SetValue(7, 4)
    
    ' Do the required sort first
    Call Arrays.Sort(myIntArray)
    
    ' Displays the values of the Array.
    Debug.Print "The int array contains the following:"
    Call PrintValues(myIntArray)
    
    ' Locates a specific object that does not exist in the Array.
    Dim myObjectOdd As Long
    myObjectOdd = 1
    Call FindMyObject(myIntArray, myObjectOdd)

    ' Locates an object that exists in the Array.
    Dim myObjectEven As Long
    myObjectEven = 6
    Call FindMyObject(myIntArray, myObjectEven)
End Sub

Private Sub FindMyObject(ByVal myArr As DotNetLib.Array, ByVal myObject As Variant)
    Dim myIndex As Long
    myIndex = Arrays.BinarySearch(myArr, myObject)
    
    If (myIndex < 0) Then
        Debug.Print VBString.Format("The object to search for ({0}) is not found. The next larger object is at index {1}.", myObject, Not myIndex)
    Else
        Debug.Print VBString.Format("The object to search for ({0}) is at index {1}.", myObject, myIndex)
    End If
End Sub

Private Sub PrintValues(ByVal myArr As DotNetLib.Array)
    Dim formatString As String
    formatString = Regex.Unescape("\t{0}")
    Dim i As Long
    Dim cols As Long
    cols = myArr.GetLength(myArr.Rank - 1)
    Dim obj As Variant
    For Each obj In myArr
        If (i < cols) Then
            i = i + 1
        Else
            Debug.Print
            i = 1
        End If
        Debug.Print VBString.Format(formatString, obj);
    Next
    Debug.Print
End Sub

' This code produces the following output.
'
' The int array contains the following:
'        2       3       6       7       8
' The object to search for (1) is not found. The next larger object is at index 0
'
' The object to search for (6) is at index 2.
