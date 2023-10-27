Attribute VB_Name = "ArrayIndexOfExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 28, 2023
'@LastModified October 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.indexof?view=netframework-4.8.1#system-array-indexof(system-array-system-object)

Option Explicit

''
' Searches for the specified object and returns the index of its first
' occurrence in a one-dimensional array.
''
Public Sub ArrayIndexOf()
    Dim stringArr As DotNetLib.Array
    Set stringArr = Arrays.CreateInitialize1D(Strings.GetType(), _
                        "the", "quick", "brown", "fox", "jumps", _
                        "over", "the", "lazy", "dog", "in", "the", _
                        "barn")

    ' Display the elements of the array.
    Debug.Print "The array contains the following values:"
    Dim i As Long
    For i = stringArr.GetLowerBound(0) To stringArr.GetUpperBound(0)
        Debug.Print Strings.Format("   [{0,2}]: {1}", i, stringArr(i))
    Next i
    
    ' Search for the first occurrence of the duplicated value.
    Dim searchString As String
    searchString = "the"
    Dim index As Long
    index = Arrays.IndexOf(stringArr, searchString)
    Debug.Print Strings.Format("The first occurrence of ""{0}"" is at index {1}.", _
                             searchString, index)

    ' Search for the first occurrence of the duplicated value in the last section of the array.
    index = Arrays.IndexOf2(stringArr, searchString, 4)
    Debug.Print Strings.Format("The first occurrence of ""{0}"" between index 4 and the end is at index {1}.", _
                                searchString, index)

    ' Search for the first occurrence of the duplicated value in a section of the array.
    Dim position As Long
    position = index + 1
    index = Arrays.IndexOf3(stringArr, searchString, position, stringArr.GetUpperBound(0) - position + 1)
    Debug.Print Strings.Format("The first occurrence of ""{0}"" between index {1} and index {2} is at index {3}.", _
                  searchString, position, stringArr.GetUpperBound(0), index)
End Sub

' The example displays the following output:
'    The array contains the following values:
'       [ 0]: the
'       [ 1]: quick
'       [ 2]: brown
'       [ 3]: fox
'       [ 4]: jumps
'       [ 5]: over
'       [ 6]: the
'       [ 7]: lazy
'       [ 8]: dog
'       [ 9]: in
'       [10]: the
'       [11]: barn
'    The first occurrence of "the" is at index 0.
'    The first occurrence of "the" between index 4 and the end is at index 6.
'    The first occurrence of "the" between index 7 and index 11 is at index 10.
