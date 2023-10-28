Attribute VB_Name = "ArrayResizeExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 28, 2023
'@LastModified October 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.resize?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how resizing affects the array.
''
Public Sub ArrayResize()
    Dim myArr As DotNetLib.Array
    Set myArr = Arrays.CreateInitialize1D(Strings.GetType(), _
                        "The", "quick", "brown", "fox", "jumps", _
                        "over", "the", "lazy", "dog")
                        
    ' Display the values of the array.
    Debug.Print "The string array initially contains the following values:"
    Call PrintIndexAndValues(myArr)
    
    ' Resize the array to a bigger size (five elements larger).
    Call Arrays.Resize(myArr, myArr.Length + 5)

    ' Display the values of the array.
    Debug.Print "After resizing to a larger size, "
    Debug.Print "the string array contains the following values:"
    Call PrintIndexAndValues(myArr)
    
    ' Resize the array to a smaller size (four elements).
    Call Arrays.Resize(myArr, 4)

    ' Display the values of the array.
    Debug.Print "After resizing to a smaller size, "
    Debug.Print "the string array contains the following values:"
    Call PrintIndexAndValues(myArr)
End Sub

Private Sub PrintIndexAndValues(ByVal myArr As DotNetLib.Array)
    Dim i As Long
    For i = 0 To myArr.Length - 1
        Debug.Print Strings.Format("   [{0}] : {1}", i, myArr(i))
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'The string array initially contains the following values:
'   [0] : The
'   [1] : quick
'   [2] : brown
'   [3] : fox
'   [4] : jumps
'   [5] : over
'   [6] : the
'   [7] : lazy
'   [8] : dog
'
'After resizing to a larger size,
'the string array contains the following values:
'   [0] : The
'   [1] : quick
'   [2] : brown
'   [3] : fox
'   [4] : jumps
'   [5] : over
'   [6] : the
'   [7] : lazy
'   [8] : dog
'   [9] :
'   [10] :
'   [11] :
'   [12] :
'   [13] :
'
'After resizing to a smaller size,
'the string array contains the following values:
'   [0] : The
'   [1] : quick
'   [2] : brown
'   [3] : fox
'
'*/
