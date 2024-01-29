Attribute VB_Name = "ArraySetValueExample"
'@Folder "Examples.System.Array.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 29, 2023
'@LastModified January 29, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.setvalue?view=netframework-4.8.1#examples
'@Dependencies ArrayEx.cls

Option Explicit

''
' The following code example demonstrates how to set and get a specific value
' in a one-dimensional or multidimensional array.
''
Public Sub ArraySetValue()
    ' Creates and initializes a one-dimensional array.
    Dim myArr1 As DotNetLib.Array
    Set myArr1 = Arrays.CreateInstance(VBString.GetType(), 5)
    
    ' Sets the element at index 3.
    Call myArr1.SetValue("three", 3)
    Debug.Print VBString.Format("[3]:   {0}", myArr1.GetValue(3))
    
    ' Creates and initializes a two-dimensional array.
    Dim myArr2 As DotNetLib.Array
    Set myArr2 = Arrays.CreateInstance2(VBString.GetType(), 5, 5)

    ' Sets the element at index 1,3.
    Call myArr2.SetValue2("one-three", 1, 3)
    Debug.Print VBString.Format("[1,3]:   {0}", myArr2.GetValue2(1, 3))

    ' Creates and initializes a three-dimensional array.
    Dim myArr3 As DotNetLib.Array
    Set myArr3 = Arrays.CreateInstance3(VBString.GetType(), 5, 5, 5)
    
    ' Sets the element at index 1,2,3.
    Call myArr3.SetValue3("one-two-three", 1, 2, 3)
    Debug.Print VBString.Format("[1,2,3]:   {0}", myArr3.GetValue3(1, 2, 3))
    
    ' Creates and initializes a seven-dimensional array.
    Dim pvtLengths() As Long
    Call VBArray.CreateInitialize1D(pvtLengths, 5, 5, 5, 5, 5, 5, 5)
    Dim myArr7 As DotNetLib.Array
    Set myArr7 = Arrays.CreateInstance4(VBString.GetType(), pvtLengths)
    
    ' Sets the element at index 1,2,3,0,1,2,3.
    Dim myIndices() As Long
    Call VBArray.CreateInitialize1D(myIndices, 1, 2, 3, 0, 1, 2, 3)
    Call myArr7.SetValue4("one-two-three-zero-one-two-three", myIndices)
    Debug.Print VBString.Format("[1,2,3,0,1,2,3]:   {0}", myArr7.GetValue4(myIndices))
    
End Sub

'/*
'This code produces the following output.
'[3]:   three
'[1,3]:   one-three
'[1,2,3]:   one-two-three
'[1,2,3,0,1,2,3]:   one-two-three-zero-one-two-three
'*/
