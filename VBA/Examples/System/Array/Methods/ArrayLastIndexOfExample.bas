Attribute VB_Name = "ArrayLastIndexOfExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 28, 2023
'@LastModified October 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.lastindexof?view=netframework-4.8.1#system-array-lastindexof(system-array-system-object)

Option Explicit

''
' The following code example shows how to determine the index of the last
' occurrence of a specified element in an array.
''
Public Sub ArrayLastIndexOf()
    Dim myArray As DotNetLib.Array
    Set myArray = Arrays.CreateInstance(VBAString.GetType(), 12)
    Call myArray.SetValue("the", 0)
    Call myArray.SetValue("quick", 1)
    Call myArray.SetValue("brown", 2)
    Call myArray.SetValue("fox", 3)
    Call myArray.SetValue("jumps", 4)
    Call myArray.SetValue("over", 5)
    Call myArray.SetValue("the", 6)
    Call myArray.SetValue("lazy", 7)
    Call myArray.SetValue("dog", 8)
    Call myArray.SetValue("in", 9)
    Call myArray.SetValue("the", 10)
    Call myArray.SetValue("barn", 11)
    
    ' Displays the values of the Array.
    Debug.Print "The Array contains the following values:"
    Call PrintIndexAndValues(myArray)
    
    ' Searches for the last occurrence of the duplicated value.
    Dim myString As String
    myString = "the"
    Dim myIndex As Long
    myIndex = Arrays.LastIndexOf(myArray, myString)
    Debug.Print VBAString.Format("The last occurrence of ""{0}"" is at index {1}.", myString, myIndex)
    
    ' Searches for the last occurrence of the duplicated value in the first section of the Array.
    myIndex = Arrays.LastIndexOf2(myArray, myString, 8)
    Debug.Print VBAString.Format("The last occurrence of ""{0}"" between the start and index 8 is at index {1}.", myString, myIndex)

    ' Searches for the last occurrence of the duplicated value in a section of the Array.
    ' Note that the start index is greater than the end index because the search is done backward.
    myIndex = Arrays.LastIndexOf3(myArray, myString, 10, 6)
    Debug.Print VBAString.Format("The last occurrence of ""{0}"" between index 5 and index 10 is at index {1}.", myString, myIndex)
End Sub

Private Sub PrintIndexAndValues(ByVal anArray As DotNetLib.Array)
    Dim formatString As String
    formatString = Regex.Unescape("\t[{0}]:\t{1}")
    Dim i As Long
    For i = anArray.GetLowerBound(0) To anArray.GetUpperBound(0)
        Debug.Print VBAString.Format(formatString, i, anArray.GetValue(i))
    Next i
End Sub

'/*
'This code produces the following output.
'
'The Array contains the following values:
'    [0]:    the
'    [1]:    quick
'    [2]:    brown
'    [3]:    fox
'    [4]:    jumps
'    [5]:    over
'    [6]:    the
'    [7]:    lazy
'    [8]:    dog
'    [9]:    in
'    [10]:   the
'    [11]:   barn
'The last occurrence of "the" is at index 10.
'The last occurrence of "the" between the start and index 8 is at index 6.
'The last occurrence of "the" between index 5 and index 10 is at index 10.
'*/
