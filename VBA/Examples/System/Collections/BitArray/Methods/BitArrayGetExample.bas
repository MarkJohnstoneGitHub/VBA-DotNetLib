Attribute VB_Name = "BitArrayGetExample"
'@Folder("Examples.System.Collections.BitArray.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 2, 2023
'@LastModified November 2, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray.get?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to set and get specific elements in a BitArray.
''
Public Sub BitArrayGetExample()
    Dim myBA As DotNetLib.BitArray
    Set myBA = BitArray.Create(5)
    
    ' Displays the properties and values of the BitArray.
    Debug.Print "myBA values:"
    Call PrintIndexAndValues(myBA)
    
    ' Sets all the elements to true.
    Call myBA.SetAll(True)
    
    ' Displays the properties and values of the BitArray.
    Debug.Print "After setting all elements to true,"
    Call PrintIndexAndValues(myBA)
    
    ' Sets the last index to false.
    Call myBA.Set(myBA.Count - 1, False)

    ' Displays the properties and values of the BitArray.
    Debug.Print "After setting the last element to false,"
    Call PrintIndexAndValues(myBA)

    ' Gets the value of the last two elements.
    Debug.Print "The last two elements are: "
    Debug.Print Strings.Format("    at index {0} : {1}", myBA.Count - 2, myBA.Get(myBA.Count - 2))
    Debug.Print Strings.Format("    at index {0} : {1}", myBA.Count - 1, myBA.Get(myBA.Count - 1))
End Sub

Private Sub PrintIndexAndValues(ByVal myCol As mscorlib.IEnumerable)
    Dim i As Long
    Dim obj As Variant
    For Each obj In myCol
         Debug.Print Strings.Format("    [{0}]:    {1}", i, obj)
         i = i + 1
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'myBA Values:
'    [0]:    False
'    [1]:    False
'    [2]:    False
'    [3]:    False
'    [4]:    False
'
'After setting all elements to true,
'    [0]:    True
'    [1]:    True
'    [2]:    True
'    [3]:    True
'    [4]:    True
'
'After setting the last element to false,
'    [0]:    True
'    [1]:    True
'    [2]:    True
'    [3]:    True
'    [4]:    False
'
'The last two elements are:
'    at index 3 : True
'    at index 4 : False
'
'*/
