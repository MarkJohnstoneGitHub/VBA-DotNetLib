Attribute VB_Name = "ArrayReverse2Example"
'@Folder "Examples.System.Array.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 28, 2023
'@LastModified October 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.reverse?view=netframework-4.8.1#system-array-reverse(system-array-system-int32-system-int32)

Option Explicit

''
' The following code example shows how to reverse the sort of the values in a
' range of elements in an Array.
''
Public Sub ArrayReverse2()
    ' Creates and initializes a new Array.
    Dim myArray As DotNetLib.Array
    Set myArray = Arrays.CreateInstance(VBString.GetType(), 9)
    Call myArray.SetValue("The", 0)
    Call myArray.SetValue("QUICK", 1)
    Call myArray.SetValue("BROWN", 2)
    Call myArray.SetValue("FOX", 3)
    Call myArray.SetValue("jumps", 4)
    Call myArray.SetValue("over", 5)
    Call myArray.SetValue("the", 6)
    Call myArray.SetValue("lazy", 7)
    Call myArray.SetValue("dog", 8)
        
    ' Displays the values of the Array.
    Debug.Print "The Array initially contains the following values:"
    Call PrintIndexAndValues(myArray)

    ' Reverses the sort of the values of the Array.
    Call Arrays.Reverse2(myArray, 1, 3)

    ' Displays the values of the Array.
    Debug.Print "After reversing:"
    Call PrintIndexAndValues(myArray)
End Sub

Private Sub PrintIndexAndValues(ByVal myArray As DotNetLib.Array)
    Dim formatString As String
    formatString = Regex.Unescape("\t[{0}]:\t{1}")
    Dim i As Long
    For i = myArray.GetLowerBound(0) To myArray.GetUpperBound(0)
        Debug.Print VBString.Format(formatString, i, myArray.GetValue(i))
    Next i
End Sub

'/*
'This code produces the following output.
'
'The Array initially contains the following values:
'    [0]:    The
'    [1]:    QUICK
'    [2]:    BROWN
'    [3]:    FOX
'    [4]:    jumps
'    [5]:    over
'    [6]:    the
'    [7]:    lazy
'    [8]:    dog
'After reversing:
'    [0]:    The
'    [1]:    FOX
'    [2]:    BROWN
'    [3]:    QUICK
'    [4]:    jumps
'    [5]:    over
'    [6]:    the
'    [7]:    lazy
'    [8]:    dog
'*/
